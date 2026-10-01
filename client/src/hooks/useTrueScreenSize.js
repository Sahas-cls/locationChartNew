import { useState, useEffect } from "react";

const useTrueScreenSize = () => {
  const [screenSize, setScreenSize] = useState({
    width: window.screen.width,
    height: window.screen.height,
  });

  useEffect(() => {
    const handleResize = () => {
      setScreenSize({
        width: window.screen.width,
        height: window.screen.height,
      });
    };
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  return screenSize;
};

// Usage in your component:
export default useTrueScreenSize;
