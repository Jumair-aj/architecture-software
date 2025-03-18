import { useState } from "react";
import * as XLSX from "xlsx";

const useExcelProcessor = () => {
  const [loading, setLoading] = useState(false);
  const [uploads, setUploads] = useState({
    upload1: { files: null, visible: true },
    upload2: { files: null, visible: true },
    upload3: { files: null, visible: true },
  });

  const handleFileChange = (id, files) => {
    setUploads((prev) => ({
      ...prev,
      [id]: {
        files: files ? Array.from(files) : null,
        visible: false,
      },
    }));
  };

  const resetUpload = (id) => {
    setUploads((prev) => ({
      ...prev,
      [id]: {
        files: null,
        visible: true,
      },
    }));
  };

  const readExcel = (file, sheetIndex = 0) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsBinaryString(file);
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: "binary" });

        // Ensure we get the correct sheet index
        const sheetName = workbook.SheetNames[sheetIndex] || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        console.log(`✅ File: ${file.name}, Rows: ${jsonData.length}`); // Debug log
        resolve(jsonData);
      };
    });
  };

  const processExcelFiles = async () => {
    const fileGroups = Object.values(uploads)
      .map((upload) => upload.files)
      .filter(Boolean);

    if (fileGroups.length < 2) {
      alert("Please upload at least two Excel files.");
      return;
    }

    setLoading(true);

    try {
      // Read files with correct sheet index
      const fileData = await Promise.all([
        readExcel(fileGroups[0][0], 1),  // First file (sheet index 1)
        readExcel(fileGroups[1][0], 0),  // Second file (sheet index 0)
        fileGroups.length > 2 ? readExcel(fileGroups[2][0], 0) : Promise.resolve([]), // Third file (optional)
      ]);

      const [epeTags, merTags, sapTags] = fileData;

      console.log("📝 Processed Data Counts:", {
        epeTags: epeTags.length,
        merTags: merTags.length,
        sapTags: sapTags.length,
      });

      if (epeTags.length === 0 || merTags.length === 0) {
        alert("Not enough data processed.");
        setLoading(false);
        return;
      }

      const merTagMap = {};
      merTags.forEach((row) => {
        merTagMap[row["MER TAG NO"]] = row;
      });

      const sapTagMap = {};
      sapTags.forEach((row) => {
        sapTagMap[row["SAP-TAGS"]] = row;
      });

      const processedData = epeTags.map((drawing, i) => {
        const tagNumber = drawing["Tag Number"];
        const merMatch = merTagMap[tagNumber];
        const sapMatch = sapTagMap[tagNumber];

        return {
          "SL.NO": i + 1,
          "Drawing Number": drawing["Drawing no"],
          "EPE Tag Number": tagNumber,
          "MER Tag No": merMatch ? merMatch["MER TAG NO"] : "NOT IN MER",
          "SITE MARKUP TAG NO": merMatch ? merMatch["SITE CHANGE"] : "",
          "SAP tag": sapMatch
            ? merMatch
              ? "NO CHANGE IN SAP"
              : "AVAILABLE IN SAP"
            : "NOT IN SAP",
          "Size - Old": merMatch ? merMatch["SIZE"] : "NOT AVAILABLE",
          "Equipment Description from SAP": sapMatch ? sapMatch["DESCRIPTION"] : "",
        };
      });

      // Add MER Tags not in EPE
      merTags.forEach((merRow) => {
        if (!epeTags.some((drawing) => drawing["Tag Number"] === merRow["MER TAG NO"])) {
          processedData.push({
            "SL.NO": processedData.length + 1,
            "Drawing Number": "NOT IN EPE",
            "EPE Tag Number": "NOT IN EPE",
            "MER Tag No": merRow["MER TAG NO"],
            "SITE MARKUP TAG NO": merRow["SITE CHANGE"] || "",
            "SAP tag": sapTagMap[merRow["MER TAG NO"]] ? "AVAILABLE IN SAP" : "NOT IN SAP",
            "Size - Old": merRow["SIZE"] || "NOT AVAILABLE",
            "Equipment Description from SAP": "",
          });
        }
      });

      // Sorting
      processedData.sort((a, b) => (a["Drawing Number"] || "").localeCompare(b["Drawing Number"] || ""));
      processedData.forEach((item, index) => {
        item["SL.NO"] = index + 1;
      });
      console.log("📝 Final Processed Data to Display:", processedData.length, processedData);
      generateExcel(processedData);
    } catch (error) {
      console.error("❌ Error processing Excel files:", error);
      alert("An error occurred while processing the files.");
    } finally {
      setLoading(false);
    }
  };

  const generateExcel = (data) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Processed Data");

    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const dataBlob = new Blob([excelBuffer], { type: "application/octet-stream" });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(dataBlob);
    link.download = "final-excel.xlsx";
    link.click();
  };

  return {
    handleFileChange,
    resetUpload,
    processExcelFiles,
    loading,
    uploads,
  };
};

export default useExcelProcessor;