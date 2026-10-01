import React, { useState, useEffect, useRef } from "react";
import axios from "axios";
import { FiRefreshCw, FiSearch } from "react-icons/fi";
import ReactPaginate from "react-paginate";
import { motion, AnimatePresence } from "framer-motion";
import concordLogo from "../assets/Concord_Logo.png";
import { HiOutlineDocumentReport } from "react-icons/hi";
import { BsFileSpreadsheet } from "react-icons/bs";

const WarehouseDashboard = () => {
  const [fabrics, setFabrics] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [search, setSearch] = useState("");
  const [pageInfo, setPageInfo] = useState({
    currentPage: 0,
    pageSize: 8,
    total: 0,
  });
  const [autoSlide, setAutoSlide] = useState(true);
  const [selectedFactory, setSelectedFactory] = useState(
    localStorage.getItem("selectedFactory") || 1,
  );
  const [reportLoading, setReportLoading] = useState(false); // Separate loading state for report

  useEffect(() => {
    if (selectedFactory) {
      fetchData();
    }
  }, [selectedFactory]);

  const slideIntervalRef = useRef(null);
  const currentPageRef = useRef(0);
  const searchRef = useRef("");
  const apiUrl = import.meta.env.VITE_API_URL;

  const handleGenerateSingleSheetReport = async () => {
    // Set report loading state only, not the main loading
    setReportLoading(true);

    try {
      const response = await axios.get(
        `${apiUrl}/generateSingleSheetReport/${selectedFactory}`,
        {
          params: {
            search: search,
          },
          responseType: "blob",
          timeout: 120000, // 2 minutes timeout
        },
      );

      // Check if response is an error blob
      if (response.data.type === "application/json") {
        const text = await response.data.text();
        const errorData = JSON.parse(text);
        throw new Error(errorData.error || "Failed to generate report");
      }

      // Create download link
      const url = window.URL.createObjectURL(
        new Blob([response.data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      );

      const link = document.createElement("a");
      link.href = url;
      link.setAttribute(
        "download",
        `fabric_stock_single_sheet_${Date.now()}.xlsx`,
      );
      document.body.appendChild(link);
      link.click();

      // Clean up
      setTimeout(() => {
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
      }, 100);
    } catch (error) {
      console.error("Single-sheet report generation failed:", error);

      // Show more specific error messages
      if (error.response) {
        if (error.response.data instanceof Blob) {
          const text = await error.response.data.text();
          try {
            const errorData = JSON.parse(text);
            alert(
              errorData.error ||
                errorData.message ||
                "Failed to generate report",
            );
          } catch {
            alert(
              `Error ${error.response.status}: ${error.response.statusText}`,
            );
          }
        } else {
          alert(
            error.response.data?.error ||
              error.response.data?.message ||
              "Failed to generate report",
          );
        }
      } else if (error.code === "ECONNABORTED") {
        alert(
          "Request timed out. The report is taking too long to generate. Please try again with fewer items.",
        );
      } else {
        alert(
          error.message ||
            "Failed to generate single-sheet report. Please try again.",
        );
      }
    } finally {
      // Set report loading to false after a small delay
      setTimeout(() => {
        setReportLoading(false);
      }, 500);
    }
  };

  // Sync refs with state
  useEffect(() => {
    currentPageRef.current = pageInfo.currentPage;
    searchRef.current = search;
  }, [pageInfo.currentPage, search]);

  const fetchData = async (page = 0, searchTerm = "") => {
    setLoading(true);
    setError(null);

    try {
      const validatedPage = Number.isInteger(page) ? Math.max(0, page) : 0;

      const response = await axios.get(`${apiUrl}/getData/${selectedFactory}`, {
        params: {
          page: validatedPage + 1,
          pageSize: pageInfo.pageSize,
          search: searchTerm,
        },
        validateStatus: (status) => status < 500,
      });

      if (!response.data?.success) {
        throw new Error(response.data?.error || "Invalid server response");
      }

      const safeData = Array.isArray(response.data.data)
        ? response.data.data.map((item) => ({
            ...item,
            TOTAL_QTY: Number(item.TOTAL_QTY) || 0,
            UOM: item.UOM || "N/A",
            BUYER_NAME:
              item.BUYER_NAME === "BL" ? "BLAKLADER" : item.BUYER_CODE || "N/A",
            MATERIAL_SUPPLIER_NAME: item.MATERIAL_SUPPLIER_NAME || "N/A",
            locations: (item.locations || []).map((loc) => ({
              FAB_PO_NO: loc.FAB_PO_NO || "N/A",
              FULL_LOCATION: loc.FULL_LOCATION || "N/A",
              INVOICE_NO: loc.INVOICE_NO || "N/A",
              QTY: Number(loc.QTY) || 0,
            })),
          }))
        : [];

      setFabrics(safeData);
      setPageInfo((prev) => ({
        ...prev,
        currentPage: validatedPage,
        total: Number(response.data.totalCount) || 0,
      }));
    } catch (err) {
      console.error("API Error:", {
        error: err,
        response: err.response?.data,
      });

      setError(
        `Failed to load data: ${err.response?.data?.error || err.message}`,
      );
      setFabrics([]);
      setPageInfo((prev) => ({ ...prev, currentPage: 0 }));

      setTimeout(() => setError(null), 5000);
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = (e) => {
    e.preventDefault();
    fetchData(0, search);
    resetAutoSlide();
  };

  const handleRefresh = () => {
    fetchData(pageInfo.currentPage, search);
    resetAutoSlide();
  };

  const handlePageClick = ({ selected }) => {
    if (!Number.isInteger(selected)) {
      console.error("Invalid page number:", selected);
      selected = 0;
    }
    fetchData(selected, search);
    resetAutoSlide();
  };

  const toggleAutoSlide = () => {
    setAutoSlide((prev) => {
      const newValue = !prev;
      if (newValue) {
        startAutoSlide();
      } else {
        clearInterval(slideIntervalRef.current);
      }
      return newValue;
    });
  };

  const startAutoSlide = () => {
    clearInterval(slideIntervalRef.current);
    slideIntervalRef.current = setInterval(() => {
      const totalPages = Math.ceil(pageInfo.total / pageInfo.pageSize);
      if (totalPages > 0) {
        console.log("change page");
        const nextPage = (currentPageRef.current + 1) % totalPages;
        fetchData(nextPage, searchRef.current);
      }
    }, 30000);
  };

  const resetAutoSlide = () => {
    if (autoSlide) {
      clearInterval(slideIntervalRef.current);
      startAutoSlide();
    }
  };

  useEffect(() => {
    fetchData(0);
    if (autoSlide) {
      startAutoSlide();
    }

    return () => {
      clearInterval(slideIntervalRef.current);
    };
  }, []);

  return (
    <div className="fixed inset-0 bg-gradient-to-bl from-blue-400 to-green-500 overflow-auto p-2">
      {/* Error Message */}
      {error && (
        <div className="fixed top-4 right-4 bg-red-500 text-white p-3 rounded-md shadow-lg z-50">
          {error}
          <button
            onClick={() => setError(null)}
            className="ml-2 text-white font-bold"
          >
            ×
          </button>
        </div>
      )}

      {/* Report Loading Overlay */}
      {reportLoading && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 backdrop-blur-sm">
          <div className="bg-white rounded-lg p-8 flex flex-col items-center shadow-2xl">
            <div className="animate-spin rounded-full h-16 w-16 border-t-4 border-b-4 border-blue-500 mb-4"></div>
            <h3 className="text-lg font-semibold text-gray-800">
              Generating Report...
            </h3>
            <p className="text-sm text-gray-600 mt-2">
              Please wait, this may take a few moments
            </p>
            <div className="w-48 h-2 bg-gray-200 rounded-full mt-4 overflow-hidden">
              <div
                className="h-full bg-blue-500 animate-pulse rounded-full"
                style={{ width: "100%" }}
              ></div>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div className="flex justify-between items-center mb-4">
        <div className="flex items-center">
          <img width="50px" src={concordLogo} alt="logoImg" />
          <div className="ml-3">
            <h1 className="font-bold text-white text-sm">
              Fabric Location Dashboard
            </h1>
            <p className="text-white/100 text-sm">
              |Powered by Concord IT Team
            </p>
          </div>
        </div>

        <form onSubmit={handleSearch} className="flex items-center gap-2">
          <label className="flex items-center cursor-pointer">
            <input
              type="checkbox"
              checked={autoSlide}
              onChange={toggleAutoSlide}
              className="form-checkbox h-5 w-5 text-blue-600 rounded focus:ring-blue-500"
            />
            <span className="ml-2 text-white text-sm">Auto Slide</span>
          </label>

          <div className="relative flex items-center justify-center">
            <input
              type="text"
              placeholder="Search fabrics..."
              className="pl-10 pr-4 py-1.5 rounded-lg"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
            <FiSearch className="absolute left-3 top-3 text-gray-500" />
          </div>

          {/* Single-sheet Report Button */}
          <button
            type="button"
            onClick={handleGenerateSingleSheetReport}
            disabled={reportLoading}
            className={`px-2 py-1 border rounded-lg shadow-md hover:bg-green-100 disabled:opacity-50 flex items-center gap-1 transition-all duration-200 ${
              reportLoading
                ? "bg-green-100 cursor-wait"
                : "bg-green-50 hover:bg-green-100"
            }`}
            title="Generate Single-Sheet Report (All data in one sheet)"
          >
            {reportLoading ? (
              <div className="animate-spin rounded-full h-5 w-5 border-t-2 border-b-2 border-green-600"></div>
            ) : (
              <BsFileSpreadsheet size={24} color="#4CAF50" />
            )}
            <span className="text-xs hidden sm:inline">
              {reportLoading ? "Generating..." : "Report"}
            </span>
          </button>

          <div className="">
            <select
              name=""
              id=""
              className="p-1.5 rounded-md outline-none"
              value={selectedFactory}
              onChange={(e) => {
                const value = e.target.value;
                setSelectedFactory(value);

                localStorage.setItem("selectedFactory", value);
              }}
            >
              <option value={1}>Concord Apparel (Pvt)Ltd</option>
              <option value={3}>Concord Manufacturing (Pvt)Ltd</option>
              <option value={4}>Guston Lanka (Pvt)Ltd</option>
              <option value={6}>M.G Apparel (Pvt)Ltd</option>
            </select>
          </div>
        </form>
      </div>

      {/* Content */}
      {loading ? (
        <div className="flex justify-center items-center h-64">
          <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-white"></div>
        </div>
      ) : (
        <>
          <AnimatePresence mode="wait">
            <motion.div
              key={pageInfo.currentPage}
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.5 }}
              className="grid grid-cols-4 gap-2 mb-4"
            >
              {fabrics.map((fabric) => (
                <motion.div
                  key={`${fabric.ITEM_CODE}-${pageInfo.currentPage}`}
                  initial={{ y: 20, opacity: 0 }}
                  animate={{ y: 0, opacity: 1 }}
                  transition={{ duration: 0.3 }}
                  className="bg-white/80 rounded-lg shadow-md overflow-hidden"
                >
                  <div className="flex flex-col h-[200px]">
                    <div className="bg-teal-600 text-white p-2 flex justify-between">
                      <h2 className="text-[10px] font-bold truncate">
                        {fabric.ITEM_CODE}
                      </h2>
                      <p className="text-xs truncate">
                        Total: {Number(fabric.TOTAL_QTY).toFixed(2)}{" "}
                        {fabric.UOM}
                      </p>
                    </div>
                    <div className="p-2 max-h-60 overflow-y-auto overflow-x-hidden flex-1">
                      <table className="w-full max-w-full overflow-x-hidden">
                        <thead className="bg-gray-100 sticky top-0">
                          <tr>
                            <th className="p-1 text-left text-[9px]">PO</th>
                            <th className="p-1 text-left text-[9px]">
                              Location
                            </th>
                            <th className="p-2 text-left text-[9px]">
                              Invoice
                            </th>
                            <th className="p-1 text-left text-[9px]">Qty</th>
                          </tr>
                        </thead>
                        <tbody>
                          {fabric.locations.map((location, i) => (
                            <tr key={i} className="border-t">
                              <td className="p-1 text-[8px]">
                                {location.FAB_PO_NO}
                              </td>
                              <td className="p-1 text-[8px]">
                                {location.FULL_LOCATION}
                              </td>
                              <td className="p-2 text-[8px]">
                                {location.INVOICE_NO || "N/A"}
                              </td>
                              <td className="p-1 text-[8px]">{location.QTY}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="bg-gray-200/50 p-2 text-[8px]">
                      <p className="truncate">
                        Supplier: {fabric.MATERIAL_SUPPLIER_NAME}
                      </p>
                      <p>
                        Buyer:{" "}
                        {fabric.BUYER_NAME === "BL"
                          ? "BLAKLADER"
                          : fabric.BUYER_NAME}
                      </p>
                    </div>
                  </div>
                </motion.div>
              ))}
            </motion.div>
          </AnimatePresence>

          {pageInfo.total > pageInfo.pageSize && (
            <div className="absolute bottom-1 left-0 right-0 flex justify-center">
              <div className="bg-white/80 backdrop-blur-sm rounded-lg p-2 opacity-90 hover:opacity-100 transition-opacity duration-300">
                <ReactPaginate
                  previousLabel={"< Previous"}
                  nextLabel={"Next >"}
                  breakLabel={"..."}
                  pageCount={Math.ceil(pageInfo.total / pageInfo.pageSize)}
                  marginPagesDisplayed={2}
                  pageRangeDisplayed={5}
                  onPageChange={handlePageClick}
                  containerClassName={"flex gap-2 items-center"}
                  pageClassName={"px-3 py-1 border rounded hover:bg-gray-100"}
                  activeClassName={"bg-blue-600 text-white"}
                  previousClassName={
                    "px-3 py-1 border rounded hover:bg-gray-100"
                  }
                  nextClassName={"px-3 py-1 border rounded hover:bg-gray-100"}
                  disabledClassName={"opacity-50 cursor-not-allowed"}
                  forcePage={pageInfo.currentPage}
                />
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default WarehouseDashboard;
