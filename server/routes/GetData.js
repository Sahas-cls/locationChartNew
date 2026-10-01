const express = require("express");
const router = express.Router();
const { poolPromise } = require("../db");
const ExcelJS = require("exceljs");

// Add this helper function for batch processing
const batchProcess = async (items, batchSize, callback) => {
  const results = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    const batchResults = await callback(batch);
    results.push(...batchResults);
  }
  return results;
};

router.get("/getData/:facId", async (req, res) => {
  const { facId } = req.params;
  let stockTable = "";
  let locationFilter = "";
  console.log("factory id: ", facId);
  switch (parseInt(facId, 10)) {
    case 1:
      stockTable = "MATERIAL_LOCATION_STOCK_ONLINE";
      locationFilter = "AND m.LOCATION_CODE = 1";
      break;

    case 3:
      stockTable = "MATERIAL_LOCATION_STOCK_CML";
      locationFilter = ""; // Add if MGA has a location filter
      break;

    case 4:
      stockTable = "MATERIAL_LOCATION_STOCK_GLL";
      locationFilter = "AND m.FACTORY_CODE = 3";
      break;

    case 6:
      stockTable = "MATERIAL_LOCATION_STOCK_MGA";
      factoryFilter = "AND COM_CODE = 6";
      break;

    default:
      return res.status(400).json({
        success: false,
        error: "Invalid factory id",
      });
  }
  console.log("requesting data", facId);
  const startTime = Date.now();

  try {
    // Validate and sanitize all inputs
    const search =
      typeof req.query.search === "string" ? req.query.search.trim() : "";
    const page = parseInt(req.query.page, 10);
    const pageSize = parseInt(req.query.pageSize, 10);

    const validatedPage = !isNaN(page) && page > 0 ? page : 1;
    const validatedPageSize =
      !isNaN(pageSize) && pageSize > 0 && pageSize <= 100 ? pageSize : 8;
    const offset = (validatedPage - 1) * validatedPageSize;

    const pool = await poolPromise;
    const request = pool.request();

    // OPTIMIZATION 1: Use CTE to avoid repeated CAST operations
    // OPTIMIZATION 2: Filter early to reduce dataset
    const itemsQuery = `
      WITH FilteredData AS (
        SELECT 
          ITEM_CODE,
          BAL_QTY,
          UOM,
          BUYER_CODE,
          MATERIAL_SUPPLIER_NAME,
          FULL_STOCK_LOCATION
        FROM ${stockTable} WITH (NOLOCK)
        WHERE BUYER_CODE LIKE 'BL'
          AND FULL_STOCK_LOCATION NOT LIKE '%CUT-END%'
          AND ISNUMERIC(BAL_QTY) = 1  -- Pre-filter invalid numeric values
          ${search ? `AND (ITEM_CODE LIKE @searchPattern OR MATERIAL_SUPPLIER_NAME LIKE @searchPattern)` : ""}
      )
      SELECT 
        ITEM_CODE,
        SUM(CAST(BAL_QTY AS FLOAT)) AS TOTAL_QTY,
        MAX(UOM) AS UOM,
        MAX(BUYER_CODE) AS BUYER_CODE,
        MAX(MATERIAL_SUPPLIER_NAME) AS MATERIAL_SUPPLIER_NAME
      FROM FilteredData
      WHERE CAST(BAL_QTY AS FLOAT) > 0
      GROUP BY ITEM_CODE
      HAVING SUM(CAST(BAL_QTY AS FLOAT)) > 0
      ORDER BY ITEM_CODE
      OFFSET @offset ROWS 
      FETCH NEXT @pageSize ROWS ONLY
    `;

    // Add search parameter with wildcards
    if (search) {
      request.input("searchPattern", `%${search}%`);
    }
    request.input("offset", offset);
    request.input("pageSize", validatedPageSize);

    const itemsResult = await request.query(itemsQuery);
    const items = itemsResult.recordset;

    if (items.length === 0) {
      return res.json({
        success: true,
        data: [],
        totalCount: 0,
        page: validatedPage,
        pageSize: validatedPageSize,
        executionTime: Date.now() - startTime,
      });
    }

    // OPTIMIZATION 3: Use temporary table for batch processing
    // Create a temp table to store item codes for better join performance
    const createTempTableQuery = `
      CREATE TABLE #TempItemCodes (ITEM_CODE NVARCHAR(100) PRIMARY KEY);
      
      ${items.map((_, i) => `INSERT INTO #TempItemCodes VALUES (@itemCode${i});`).join("\n")}
      
      SELECT
        m.ITEM_CODE,
        m.FAB_PO_NO,
        m.INVOICE_NO,
        m.FULL_STOCK_LOCATION,
        SUM(CAST(m.BAL_QTY AS FLOAT)) AS QTY
      FROM ${stockTable} m WITH (NOLOCK)
      INNER JOIN #TempItemCodes t ON m.ITEM_CODE = t.ITEM_CODE
      WHERE CAST(m.BAL_QTY AS FLOAT) > 0
        AND m.LOCATION_CODE = 1
        AND m.FULL_STOCK_LOCATION NOT LIKE '%CUT-END%'
        AND ISNUMERIC(m.BAL_QTY) = 1
      GROUP BY
        m.ITEM_CODE,
        m.FAB_PO_NO,
        m.INVOICE_NO,
        m.FULL_STOCK_LOCATION
      ORDER BY m.ITEM_CODE, m.FAB_PO_NO;
      
      DROP TABLE #TempItemCodes;
    `;

    // Add all item codes as parameters
    items.forEach((item, i) => {
      request.input(`itemCode${i}`, item.ITEM_CODE);
    });

    const locationsResult = await request.query(createTempTableQuery);
    // Extract the result from multiple recordset returns
    const allLocations = locationsResult.recordset || [];

    // OPTIMIZATION 4: Create a Map for O(1) lookup instead of O(n) filter
    const locationsMap = new Map();
    allLocations.forEach((loc) => {
      if (!locationsMap.has(loc.ITEM_CODE)) {
        locationsMap.set(loc.ITEM_CODE, []);
      }
      locationsMap.get(loc.ITEM_CODE).push({
        FAB_PO_NO: loc.FAB_PO_NO || "N/A",
        FULL_LOCATION: loc.FULL_STOCK_LOCATION || "N/A",
        INVOICE_NO: loc.INVOICE_NO || "N/A",
        QTY: parseFloat(loc.QTY) || 0,
      });
    });

    // Format response using Map for better performance
    const formattedData = items.map((item) => ({
      ITEM_CODE: item.ITEM_CODE || "N/A",
      TOTAL_QTY: parseFloat(item.TOTAL_QTY) || 0,
      UOM: item.UOM || "N/A",
      BUYER_NAME: item.BUYER_CODE || "N/A",
      MATERIAL_SUPPLIER_NAME: item.MATERIAL_SUPPLIER_NAME || "N/A",
      locations: locationsMap.get(item.ITEM_CODE) || [],
    }));

    // OPTIMIZATION 5: Optimized count query with indexes
    const countQuery = `
      SELECT COUNT(DISTINCT ITEM_CODE) as total
      FROM ${stockTable} WITH (NOLOCK)
      WHERE BUYER_CODE LIKE 'BL'
        AND FULL_STOCK_LOCATION NOT LIKE '%CUT-END%'
        AND ISNUMERIC(BAL_QTY) = 1
        AND CAST(BAL_QTY AS FLOAT) > 0
        ${search ? `AND (ITEM_CODE LIKE @searchPattern OR MATERIAL_SUPPLIER_NAME LIKE @searchPattern)` : ""}
    `;

    // Reuse the same search parameter if exists
    if (search && !request.parameters.searchPattern) {
      request.input("searchPattern", `%${search}%`);
    }

    const countResult = await request.query(countQuery);

    console.log(`Query executed in ${Date.now() - startTime}ms`);

    res.json({
      success: true,
      data: formattedData,
      totalCount: countResult.recordset[0].total,
      page: validatedPage,
      pageSize: validatedPageSize,
      executionTime: Date.now() - startTime,
    });
  } catch (error) {
    console.error("Database error:", {
      message: error.message,
      stack: error.stack,
      query: req.query,
      time: new Date().toISOString(),
    });

    // Check for timeout specifically
    if (error.message && error.message.includes("Timeout")) {
      return res.status(504).json({
        success: false,
        error:
          "Query timeout - please try with smaller page size or refine search criteria",
        details:
          process.env.NODE_ENV === "development"
            ? error.message
            : "Request timeout",
      });
    }

    res.status(500).json({
      success: false,
      error: "Database operation failed",
      details:
        process.env.NODE_ENV === "development"
          ? error.message
          : "Internal server error",
      executionTime: Date.now() - startTime,
    });
  }
});

router.get("/generateReport", async (req, res) => {
  console.log("Generating report...");
  try {
    // Get all data without pagination
    const search = req.query.search || "";
    const pool = await poolPromise;
    const request = pool.request();

    // Get all items with their total quantities
    const itemsQuery = `
      SELECT
        ITEM_CODE,
        SUM(BAL_QTY) AS TOTAL_QTY,
        MAX(UOM) AS UOM,
        MAX(BUYER_CODE) AS BUYER_CODE,
        MAX(MATERIAL_SUPPLIER_NAME) AS MATERIAL_SUPPLIER_NAME
      FROM MATERIAL_LOCATION_STOCK_ONLINE
      WHERE BUYER_CODE LIKE 'BL' AND BAL_QTY > 0
        ${search ? `AND (ITEM_CODE LIKE '%' + @search + '%' OR MATERIAL_SUPPLIER_NAME LIKE '%' + @search + '%')` : ""}
      GROUP BY ITEM_CODE
      ORDER BY ITEM_CODE
    `;

    if (search) request.input("search", search);
    const itemsResult = await request.query(itemsQuery);
    const items = itemsResult.recordset;

    if (items.length === 0) {
      return res.status(404).json({
        success: false,
        message: "No data found to generate report",
      });
    }

    // Get all locations for these items
    const paramNames = items.map((_, i) => `@itemCode${i}`).join(",");
    const locationsQuery = `
      SELECT
        ITEM_CODE,
        FAB_PO_NO,
        BAL_QTY,
        INVOICE_NO,
        FULL_STOCK_LOCATION,
        SUM(BAL_QTY) AS QTY
      FROM MATERIAL_LOCATION_STOCK_ONLINE
      WHERE ITEM_CODE IN (${paramNames})
      AND BAL_QTY > 0
      GROUP BY
        ITEM_CODE,
        BAL_QTY,
        FAB_PO_NO,
        INVOICE_NO,
        FULL_STOCK_LOCATION
      ORDER BY ITEM_CODE, FAB_PO_NO
    `;

    items.forEach((item, i) => {
      request.input(`itemCode${i}`, item.ITEM_CODE);
    });

    const locationsResult = await request.query(locationsQuery);
    const allLocations = locationsResult.recordset;

    // Create Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Concord IT Team";
    workbook.created = new Date();

    // Create a worksheet for each fabric
    items.forEach((item) => {
      const locations = allLocations
        .filter((loc) => loc.ITEM_CODE === item.ITEM_CODE)
        .map((loc) => ({
          FAB_PO_NO: loc.FAB_PO_NO || "N/A",
          FULL_LOCATION: loc.FULL_STOCK_LOCATION || "N/A",
          INVOICE_NO: loc.INVOICE_NO || "N/A",
          QTY: parseFloat(loc.QTY) || 0,
        }));

      // Create sheet for this fabric
      const sheetName = item.ITEM_CODE.substring(0, 31); // Excel sheet name max 31 chars
      const worksheet = workbook.addWorksheet(sheetName);

      // Add header with fabric details
      worksheet.addRow(["Fabric Code:", item.ITEM_CODE]);
      worksheet.addRow(["Total Quantity:", item.TOTAL_QTY, item.UOM]);
      worksheet.addRow(["Supplier:", item.MATERIAL_SUPPLIER_NAME]);
      worksheet.addRow([
        "Buyer:",
        item.BUYER_CODE === "BL" ? "BLAKLADER" : item.BUYER_CODE,
      ]);
      worksheet.addRow([]); // Empty row for spacing

      // Add location details header
      const headerRow = worksheet.addRow([
        "PO Number",
        "Location",
        "Invoice No",
        "Quantity",
      ]);

      // Style the header
      headerRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF4CAF50" },
        };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      // Add location data
      locations.forEach((loc) => {
        const row = worksheet.addRow([
          loc.FAB_PO_NO,
          loc.FULL_LOCATION,
          loc.INVOICE_NO,
          loc.QTY,
        ]);

        // Style data rows
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        });
      });

      // Auto-fit columns
      worksheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          const columnLength = cell.value ? cell.value.toString().length : 10;
          if (columnLength > maxLength) {
            maxLength = columnLength;
          }
        });
        column.width = Math.min(Math.max(maxLength + 2, 15), 50);
      });

      // Add total row
      const totalRow = worksheet.addRow([
        "TOTAL",
        "",
        "",
        locations.reduce((sum, loc) => sum + loc.QTY, 0),
      ]);

      totalRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFE8F5E9" },
        };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Add a summary sheet
    const summarySheet = workbook.addWorksheet("Summary");
    summarySheet.addRow(["Fabric Stock Summary Report"]);
    summarySheet.addRow(["Generated:", new Date().toLocaleString()]);
    summarySheet.addRow([]);

    const summaryHeader = summarySheet.addRow([
      "S.No",
      "Fabric Code",
      "Total Quantity",
      "UOM",
      "Supplier",
      "Buyer",
      "Number of Locations",
    ]);

    summaryHeader.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2196F3" },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    });

    items.forEach((item, index) => {
      const locations = allLocations.filter(
        (loc) => loc.ITEM_CODE === item.ITEM_CODE,
      );
      const row = summarySheet.addRow([
        index + 1,
        item.ITEM_CODE,
        item.TOTAL_QTY,
        item.UOM,
        item.MATERIAL_SUPPLIER_NAME,
        item.BUYER_CODE === "BL" ? "BLAKLADER" : item.BUYER_CODE,
        locations.length,
      ]);

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Auto-fit summary columns
    summarySheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = Math.min(Math.max(maxLength + 2, 15), 50);
    });

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Set response headers
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=fabric_stock_report_${Date.now()}.xlsx`,
    );
    res.setHeader("Content-Length", buffer.length);

    // Send the file
    res.send(buffer);
  } catch (error) {
    console.error("Report generation failed:", error);
    res.status(500).json({
      success: false,
      error: "Report generation failed",
      details: error.message,
    });
  }
});

router.get("/generateSingleSheetReport/:facId", async (req, res) => {
  console.log("Generating single sheet report...");
  const { facId } = req.params;
  const factoryId = parseInt(facId, 10);

  let stockTable = "";
  let factoryFilter = "";

  // Match the factory mapping from getData endpoint
  switch (factoryId) {
    case 1:
      stockTable = "MATERIAL_LOCATION_STOCK_ONLINE";
      factoryFilter = "AND LOCATION_CODE = 1";
      break;

    case 3:
      stockTable = "MATERIAL_LOCATION_STOCK_CML";
      factoryFilter = "";
      break;

    case 4:
      stockTable = "MATERIAL_LOCATION_STOCK_GLL";
      factoryFilter = "AND COM_CODE = 4";
      break;

    case 6:
      stockTable = "MATERIAL_LOCATION_STOCK_MGA";
      factoryFilter = "AND COM_CODE = 6";
      break;

    default:
      return res.status(400).json({
        success: false,
        error: "Invalid factory id",
      });
  }

  try {
    const search = req.query.search || "";
    const pool = await poolPromise;
    const request = pool.request();

    // Get all items with their total quantities - matching getData logic
    const itemsQuery = `
      SELECT
        ITEM_CODE,
        SUM(CAST(BAL_QTY AS FLOAT)) AS TOTAL_QTY,
        MAX(UOM) AS UOM,
        MAX(BUYER_CODE) AS BUYER_CODE,
        MAX(MATERIAL_SUPPLIER_NAME) AS MATERIAL_SUPPLIER_NAME
      FROM ${stockTable} WITH (NOLOCK)
      WHERE BUYER_CODE LIKE 'BL' 
        AND FULL_STOCK_LOCATION NOT LIKE '%CUT-END%'
        AND ISNUMERIC(BAL_QTY) = 1
        AND CAST(BAL_QTY AS FLOAT) > 0
        ${factoryFilter}
        ${search ? `AND (ITEM_CODE LIKE '%' + @search + '%' OR MATERIAL_SUPPLIER_NAME LIKE '%' + @search + '%')` : ""}
      GROUP BY ITEM_CODE
      HAVING SUM(CAST(BAL_QTY AS FLOAT)) > 0
      ORDER BY ITEM_CODE
    `;

    if (search) request.input("search", search);
    const itemsResult = await request.query(itemsQuery);
    const items = itemsResult.recordset;

    if (items.length === 0) {
      return res.status(404).json({
        success: false,
        message: "No data found to generate report",
      });
    }

    // Get all locations for these items - matching getData logic
    const paramNames = items.map((_, i) => `@itemCode${i}`).join(",");
    const locationsQuery = `
      SELECT
        ITEM_CODE,
        FAB_PO_NO,
        INVOICE_NO,
        FULL_STOCK_LOCATION,
        SUM(CAST(BAL_QTY AS FLOAT)) AS QTY
      FROM ${stockTable} WITH (NOLOCK)
      WHERE ITEM_CODE IN (${paramNames})
        AND CAST(BAL_QTY AS FLOAT) > 0
        AND FULL_STOCK_LOCATION NOT LIKE '%CUT-END%'
        AND ISNUMERIC(BAL_QTY) = 1
        ${factoryFilter}
      GROUP BY
        ITEM_CODE,
        FAB_PO_NO,
        INVOICE_NO,
        FULL_STOCK_LOCATION
      ORDER BY ITEM_CODE, FAB_PO_NO
    `;

    items.forEach((item, i) => {
      request.input(`itemCode${i}`, item.ITEM_CODE);
    });

    const locationsResult = await request.query(locationsQuery);
    const allLocations = locationsResult.recordset;

    // Create Excel workbook with single sheet
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Concord IT Team";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Fabric Stock Report");

    // Add title
    worksheet.addRow(["FABRIC STOCK REPORT"]);
    worksheet.addRow([`Generated: ${new Date().toLocaleString()}`]);
    worksheet.addRow([`Filter: ${search || "All Fabrics"}`]);
    worksheet.addRow([]); // Empty row

    // Merge title cells for better appearance
    worksheet.mergeCells("A1:F1");
    worksheet.getCell("A1").font = { bold: true, size: 16 };
    worksheet.getCell("A1").alignment = { horizontal: "center" };

    // Main headers
    const headers = [
      "Fabric Code",
      "Total Qty",
      "UOM",
      "Supplier",
      "Buyer",
      "PO Number",
      "Location",
      "Invoice No",
      "Location Qty",
    ];

    const headerRow = worksheet.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4CAF50" },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Add data
    items.forEach((item) => {
      const locations = allLocations.filter(
        (loc) => loc.ITEM_CODE === item.ITEM_CODE,
      );

      locations.forEach((loc, index) => {
        const rowData = [
          index === 0 ? item.ITEM_CODE : "", // Show fabric code only once
          index === 0 ? item.TOTAL_QTY : "",
          index === 0 ? item.UOM : "",
          index === 0 ? item.MATERIAL_SUPPLIER_NAME : "",
          index === 0
            ? item.BUYER_CODE === "BL"
              ? "BLAKLADER"
              : item.BUYER_CODE
            : "",
          loc.FAB_PO_NO || "N/A",
          loc.FULL_STOCK_LOCATION || "N/A",
          loc.INVOICE_NO || "N/A",
          parseFloat(loc.QTY) || 0,
        ];

        const row = worksheet.addRow(rowData);

        // Style data rows
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        });

        // Highlight fabric code cells
        if (index === 0) {
          row.getCell(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF5F5F5" },
          };
          row.getCell(1).font = { bold: true };
        }
      });

      // Add a blank row after each fabric for readability
      if (locations.length > 0) {
        worksheet.addRow([]);
      }
    });

    // Add summary at the bottom
    worksheet.addRow([]);
    const summaryRow = worksheet.addRow(["SUMMARY"]);
    worksheet.mergeCells(`A${worksheet.rowCount}:I${worksheet.rowCount}`);
    summaryRow.getCell(1).font = { bold: true, size: 12 };
    summaryRow.getCell(1).alignment = { horizontal: "center" };

    // Summary headers
    worksheet.addRow([
      "Fabric Code",
      "Total Qty",
      "UOM",
      "Number of Locations",
    ]);
    const summaryHeaderRow = worksheet.getRow(worksheet.rowCount);
    summaryHeaderRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2196F3" },
      };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Add summary data
    items.forEach((item) => {
      const locations = allLocations.filter(
        (loc) => loc.ITEM_CODE === item.ITEM_CODE,
      );
      worksheet.addRow([
        item.ITEM_CODE,
        item.TOTAL_QTY,
        item.UOM,
        locations.length,
      ]);
    });

    // Auto-fit all columns
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = Math.min(Math.max(maxLength + 2, 15), 50);
    });

    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();

    // Set response headers
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=fabric_stock_single_sheet_${Date.now()}.xlsx`,
    );
    res.setHeader("Content-Length", buffer.length);

    // Send the file
    res.send(buffer);
  } catch (error) {
    console.error("Report generation failed:", error);
    res.status(500).json({
      success: false,
      error: "Report generation failed",
      details: error.message,
    });
  }
});

module.exports = router;
