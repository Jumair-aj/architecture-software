import { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

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
        const sheetName =
          workbook.SheetNames[sheetIndex] || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        // console.log(`✅ File: ${file.name}, Rows: ${jsonData.length}`); // Debug log
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
        readExcel(fileGroups[0][0], 1), // First file (sheet index 1)
        readExcel(fileGroups[fileGroups.length - 1][0], 0), // Third file (optional)
      ]);

      readExcel(fileGroups[1][0], 0); // Second file (sheet index 0)
      let merTags = [];
      await Promise.all(
        fileGroups.slice(1, -1).map(async (file) => {
          const fileData = await readExcel(file[0], 0); // Assuming sheet index 0 for all file2 files
          // console.log(file[0].name.split('.')[0].split('-').join(''))
          const fileDataWithFilename = fileData.map((row) => ({
            ...row,
            filename: file[0].name.split(".")[0].split("-").join(""), // Add the filename property to each row
          }));
          merTags = merTags.concat(fileDataWithFilename);
        })
      );
      const [epeTags, sapTags] = fileData;
      const Dname = fileGroups.map((file) =>
        file[0].name.split(".")[0].split("-").join("")
      );

      // console.log("📝 Processed Data Counts:", {
      //   epeTags: epeTags.length,
      //   merTags: merTags.length,
      //   sapTags: sapTags.length,
      // });

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
        // console.log(sapMatch ? sapMatch["DESCRIPTION"] : "")
        let dngNo = "";
        if (
          drawing["Drawing no"].split("-").length > 3 ||
          drawing["Drawing no"].length != 14
        ) {
          dngNo = drawing["Drawing no"];
        } else {
          dngNo =
            drawing["Drawing no"].slice(0, 4) +
            "-" +
            drawing["Drawing no"].slice(4, 5) +
            "-" +
            drawing["Drawing no"].slice(5, 7) +
            "-" +
            drawing["Drawing no"].slice(7, 11) +
            "-" +
            drawing["Drawing no"].slice(11);
        }

        return {
          "SL.NO": i + 1,
          "Drawing Number": dngNo || "",
          "EPE Tag Number": tagNumber || "",
          "MER Tag No": merMatch
            ? merMatch["MER TAG NO"]
            : Dname.includes(drawing["Drawing no"])
            ? "NOT IN MER"
            : "",
          "Site Markup Tag No": merMatch ? merMatch["SITE CHANGE"] : "",
          "Final MER Tag No": merMatch
            ? merMatch["SITE CHANGE"] && merMatch["SITE CHANGE"].includes("NO")
              ? merMatch["MER TAG NO"]
              : merMatch["SITE CHANGE"]
              ? merMatch["SITE CHANGE"]
              : merMatch["MER TAG NO"]
            : "NOT IN MER",
          "SAP tag": sapMatch
            ? merMatch
              ? "NO CHANGE IN SAP"
              : "AVAILABLE IN SAP"
            : "NOT IN SAP",
          MERRemarks: "",
          "Equipment Description from SAP": sapMatch
            ? sapMatch["DESCRIPTION"]
            : "",
          "Equipment Type - New": "",
          "Size - Old": merMatch
            ? merMatch["Size - Old"]
              ? merMatch["Size - Old"]
              : merMatch["Size - New"]
              ? ""
              : "NOT AVAILABLE"
            : "NOT AVAILABLE",
          "Size From SAP": "",
          "Size - New": merMatch ? merMatch["Size - New"] : "",
          "Drawing No.": dngNo || drawing["Drawing no"] || "",
          Rev: "",
          "PCR / Project No.": merMatch ? merMatch["PCR / Project No."] : "",
          "Additional Information": "",
          "DRAWING LINK": "",
          "ECM LINK": "",
          "OAO LINK": "",
        };
      });

      // Add MER Tags not in EPE
      merTags.forEach((merRow) => {
        if (
          !epeTags.some(
            (drawing) => drawing["Tag Number"] === merRow["MER TAG NO"]
          )
        ) {
          let dngNo = "";
          // console.log("asd", merRow.filename)
          if (
            merRow.filename.split("-").length > 3 ||
            merRow.filename.length != 14
          ) {
            dngNo = merRow.filename;
          } else {
            dngNo =
              merRow.filename.slice(0, 4) +
              "-" +
              merRow.filename.slice(4, 5) +
              "-" +
              merRow.filename.slice(5, 7) +
              "-" +
              merRow.filename.slice(7, 11) +
              "-" +
              merRow.filename.slice(11);
          }
          processedData.push({
            "SL.NO": processedData.length + 1,
            "Drawing Number": dngNo,
            "EPE Tag Number": "NOT IN EPE",
            "MER Tag No": merRow["MER TAG NO"] || "",
            "Site Markup Tag No": merRow ? merRow["SITE CHANGE"] : "",
            "Final MER Tag No": merRow
              ? merRow["SITE CHANGE"] && merRow["SITE CHANGE"].includes("NO")
                ? merRow["MER TAG NO"]
                : merRow["SITE CHANGE"]
                ? merRow["SITE CHANGE"]
                : merRow["MER TAG NO"]
              : "",
            "SAP tag": "NOT IN SAP",
            MERRemarks: "",
            "Equipment Description from SAP": "",
            "Equipment Type - New": "",
            "Size - Old": merRow
              ? merRow["Size - Old"] || merRow["Size - New"]
                ? ""
                : "NOT AVAILABLE"
              : "",
            "Size From SAP": "",
            "Size - New": merRow ? merRow["Size - New"] : "",
            "Drawing No.": dngNo || "",
            Rev: "",
            "PCR / Project No.": merRow ? merRow["PCR / Project No."] : "",
            "Additional Information": "",
            "DRAWING LINK": "",
            "ECM LINK": "",
            "OAO LINK": "",
          });
        }
      });

      // Sorting
      processedData.sort((a, b) =>
        (a["Drawing Number"] || "").localeCompare(b["Drawing Number"] || "")
      );
      processedData.forEach((item, index) => {
        item["SL.NO"] = index + 1;
      });
      // console.log("📝 Final Processed Data to Display:", processedData.length, processedData);
      generateExcel(processedData);
    } catch (error) {
      console.error("❌ Error processing Excel files:", error);
      alert("An error occurred while processing the files.");
    } finally {
      setLoading(false);
    }
  };

  const generateExcel = async (data) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Architect");

    // Set row height for header row
    worksheet.getRow(1).height = 50;

    // Define column headers and styles
    worksheet.columns = [
      { header: "SL.NO", key: "SL.NO", width: 10 },
      { header: "Drawing Number", key: "Drawing Number", width: 20 },
      { header: "EPE Tag Number", key: "EPE Tag Number", width: 20 },
      { header: "MER Tag No", key: "MER Tag No", width: 20 },
      { header: "Site Markup Tag No", key: "Site Markup Tag No", width: 20 },
      { header: "Final MER Tag No", key: "Final MER Tag No", width: 20 },
      { header: "SAP tag", key: "SAP tag", width: 20 },
      {
        header: "Equipment Description from SAP",
        key: "Equipment Description from SAP",
        width: 30,
      },
      {
        header: "Equipment Type - New",
        key: "Equipment Type - New",
        width: 20,
      },
      { header: "Size - Old", key: "Size - Old", width: 20 },
      { header: "Size From SAP", key: "Size From SAP", width: 20 },
      { header: "Size - New", key: "Size - New", width: 20 },
      { header: "Drawing No.", key: "Drawing No.", width: 20 },
      { header: "Rev", key: "Rev", width: 10 },
      { header: "PCR / Project No.", key: "PCR / Project No.", width: 20 },
      {
        header: "Additional Information",
        key: "Additional Information",
        width: 30,
      },
      { header: "DRAWING LINK", key: "DRAWING LINK", width: 20 },
      { header: "ECM LINK", key: "ECM LINK", width: 20 },
      { header: "OAO LINK", key: "OAO LINK", width: 20 },
    ];
    // Add data rows
    worksheet.addRows(data);

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "000000" }, size: 11 };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "C0C0C0" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };

    // Apply specific cell background color for given columns
    const highlightColumns = [3, 5, 8, 11, 17, 18, 19]; // Columns that need highlighting
    highlightColumns.forEach((colIndex) => {
      headerRow.getCell(colIndex).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFC000" },
      };
    });

    // Adjust styles for all rows & apply borders
    worksheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Auto-fit columns based on content
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        maxLength = Math.max(maxLength, columnLength);
      });
      column.width = maxLength + 5; // Add some padding
    });

    // Generate the Excel file and trigger download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "final-excel.xlsx";
    link.click();
  };

  const processAndMergeSldExcelFiles = async () => {
    const fileGroups = Object.values(uploads)
      .map((upload) => upload.files)
      .filter(Boolean);

    if (fileGroups.length < 3) {
      alert("Please upload all three Excel files.");
      return;
    }

    setLoading(true);

    try {
      const fileData = await Promise.all([
        readExcel(fileGroups[0][0], 0),
        readExcel(fileGroups[1][0], 0),
        readExcel(fileGroups[2][0], 0),
      ]);

      const [epeTags, merTags, sapTags] = fileData;

      console.log("📝 Data Counts:", {
        epeTags: epeTags.length,
        merTags: merTags.length,
        sapTags: sapTags.length,
      });

      const data = epeTags.map((drawing, i) => {
        const matchingTagWithMer = merTags.find(
          (tag) =>
            tag["MER Tag Number"] === drawing["EPE Tag Number"] &&
            tag["Drawing Number"] === drawing["Drawing Number"]
        );
        const matchingTagWithSap = sapTags.find(
          (tag) => tag["SAP TAG "] === drawing["EPE Tag Number"]
        );
        console.log("EPE:", drawing?.["Size Old"], "MER:", matchingTagWithMer?.["Size Old"]); 
        return {
          "SL.NO": i + 1,
          "Drawing Number": drawing["Drawing Number"] || "",
          "EPE Tag Number": drawing["EPE Tag Number"] || "",
          "MER Tag No": matchingTagWithMer
            ? matchingTagWithMer["MER Tag Number"]
            : "NOT IN MER",
          "Site Markup Tag No": matchingTagWithMer
            ? matchingTagWithMer["site chainges"] || "NO CHANGE IN SITE"
            : "NO CHANGE IN SITE",
          "SAP tag": matchingTagWithSap
            ? matchingTagWithMer
              ? "NO CHANGE IN SAP"
              : "AVAILABLE IN SAP"
            : "NOT IN SAP",
          "Equipment Description From SAP": matchingTagWithSap
            ? matchingTagWithSap["SAP DISCRIPTION"] || ""
            : "",
          "Equipment Type - New": matchingTagWithMer
            ? matchingTagWithMer["Equipment Type-New"] || ""
            : "",
            "Size Old":
            matchingTagWithMer?.["Size Old"] && matchingTagWithMer["Size Old"].trim() !== ""
              ? matchingTagWithMer["Size Old"]
              : drawing?.["Size Old"]?.trim() !== ""
              ? drawing["Size Old"]
              : "",
          // "Size - Old": matchingTagWithMer
          //   ? matchingTagWithMer["Size - Old"]
          //     ? matchingTagWithMer["Size - Old"]
          //     : matchingTagWithMer["size new"]
          //     ? ""
          //     : "NOT AVAILABLE"
          //   : "NOT AVAILABLE",
          "Size - New": matchingTagWithMer
            ? matchingTagWithMer["size new"] || ""
            : "",
          "Drawing No.": drawing["Drawing Number"] || "",
          Rev: "",
          "PCR / Project No.": matchingTagWithMer
            ? matchingTagWithMer["PCR/PROJECT"] || ""
            : "",
          "Additional Information": matchingTagWithMer
            ? matchingTagWithMer["REMARK"] || ""
            : "",
          "DRAWING LINK": "",
          "ECM LINK": "",
          "OAO LINK": "",
        };
      });

      // Add MER Tags not in EPE
      merTags.forEach((tag) => {
        if (
          !epeTags.some(
            (drawing) =>
              drawing["EPE Tag Number"] === tag["MER Tag Number"] &&
              tag["Drawing Number"] === drawing["Drawing Number"]
          )
        ) {
          const matchingTagWithSap = sapTags.find(
            (sap) => sap["SAP TAG "] === tag["MER Tag Number"]
          );
          data.push({
            "SL.NO": data.length + 1,
            "Drawing Number": tag["Drawing Number"],
            "EPE Tag Number": "NOT IN EPE",
            "MER Tag No": tag["MER Tag Number"] || "NOT IN MER",
            "Site Markup Tag No": tag["site chainges"] || "NO CHANGE IN SITE",
            "SAP tag": matchingTagWithSap ? "AVAILABLE IN SAP" : "NOT IN SAP",
            "Equipment Description From SAP": matchingTagWithSap
              ? matchingTagWithSap["SAP DISCRIPTION"] || ""
              : "",
            "Equipment Type - New": tag ? tag["Equipment Type-New"] || "" : "",
           "Size Old": tag?.["Size Old"] ?? "",
            // "Size - Old": tag
            //   ? tag["Size - Old"]
            //     ? tag["Size - Old"]
            //     : tag["size new"]
            //     ? ""
            //     : "NOT AVAILABLE"
            //   : "NOT AVAILABLE",
            "Size - New": tag ? tag["size new"] || "" : "",
            "Drawing No.": tag["Drawing Number"] || "",
            s: "",
            "PCR / Project No.": tag ? tag["PCR/PROJECT"] || "" : "",
            "Additional Information": tag ? tag["REMARK"] || "" : "",
            "DRAWING LINK": "",
            "ECM LINK": "",
            "OAO LINK": "",
          });
        }
      });

      // 🔹 Sorting based on Drawing Number
      data.sort((a, b) =>
        (a["Drawing Number"] || "").localeCompare(b["Drawing Number"] || "")
      );

      console.log("✅ Final Processed Data:", data.length);

      generateMergedExcel(data);
    } catch (error) {
      console.error("❌ Error processing files:", error);
      alert("An error occurred while processing.");
    } finally {
      setLoading(false);
    }
  };

  // 🔹 New function to generate and download the Excel file
  const generateMergedExcel = async (data) => {
    // Create a new workbook and worksheet using ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Architect");
    worksheet.getRow(1).height = 30; // Set height for row 1 to 30 units

    // Set the height for multiple rows in a loop
    for (let i = 1; i <= worksheet.rowCount; i++) {
      worksheet.getRow(i).height = 25; // Set height for all rows to 25 units
    }

    // Define the columns for the worksheet
    worksheet.columns = [
      { header: "SL.NO", key: "SL.NO", width: 10 },
      { header: "Drawing Number", key: "Drawing Number", width: 20 },
      { header: "EPE Tag Number", key: "EPE Tag Number", width: 20 },
      { header: "MER Tag No", key: "MER Tag No", width: 20 },
      { header: "Site Markup Tag No", key: "Site Markup Tag No", width: 20 },
      { header: "SAP tag", key: "SAP tag", width: 20 },
      {
        header: "Equipment Description From SAP",
        key: "Equipment Description From SAP",
        width: 30,
      },
      {
        header: "Equipment Type - New",
        key: "Equipment Type - New",
        width: 20,
      },
      { header: "Size Old", key: "Size Old", width: 20 },
      // { header: "Size - Old", key: "Size - Old", width: 20 },
      { header: "Size - New", key: "Size - New", width: 20 },
      { header: "Drawing No.", key: "Drawing No.", width: 20 },
      { header: "Rev", key: "Rev", width: 10 },
      { header: "PCR / Project No.", key: "PCR / Project No.", width: 20 },
      {
        header: "Additional Information",
        key: "Additional Information",
        width: 30,
      },
      { header: "DRAWING LINK", key: "DRAWING LINK", width: 20 },
      { header: "ECM LINK", key: "ECM LINK", width: 20 },
      { header: "OAO LINK", key: "OAO LINK", width: 20 },
    ];

    // Add data rows to the worksheet
    worksheet.addRows(data);

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 12 };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "4472C4" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };

    // Set height for the header row
    worksheet.getRow(1).height = 30;

    // Adjust row height for all other rows
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber > 1) {
        row.height = 50;
      }
      row.eachCell((cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    });

    // Auto-fit columns based on content
    worksheet.columns.forEach((column) => {
      let maxLength = 0;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength + 2; // Add some padding
    });

    // Center-align all cells in the worksheet
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
    });

    const sortAndMergeColumns = () => {
      const EPE_COL_INDEX = 3;
      const MER_COL_INDEX = 4;
      const SL_NO_COL_INDEX = 1;
      const DRAWING_NUMBER_COL_INDEX = 2; // Adjust this if the Drawing Number is in a different column
      const columnsToMerge = [1, 5, 6, 7, 8, 9, 10]; // SL.NO, Site Markup Tag No, SAP tag, Equipment Description From SAP, Size - Old, Size - New

      // Function to safely get cell value
      const getCellValue = (row, col) => {
        try {
          return worksheet.getCell(row, col).value;
        } catch (error) {
          console.error(
            `Error getting cell value at row ${row}, column ${col}: ${error.message}`
          );
          return null;
        }
      };

      // Function to safely set cell value
      const setCellValue = (row, col, value) => {
        try {
          worksheet.getCell(row, col).value = value;
        } catch (error) {
          console.error(
            `Error setting cell value at row ${row}, column ${col}: ${error.message}`
          );
        }
      };

      // Step 1: Collect all row data and merge information
      const rowData = [];
      const mergeInfo = new Map();
      for (let i = 2; i <= worksheet.rowCount; i++) {
        const row = {};
        for (let j = 1; j <= worksheet.columnCount; j++) {
          const cell = worksheet.getCell(i, j);
          row[j] = cell.value;
          if (cell.isMerged) {
            const master = cell.master;
            if (master.address === cell.address) {
              mergeInfo.set(`${i},${j}`, {
                start: { row: master.row, col: master.col },
                end: {
                  row: master.row + master.rowSpan - 1,
                  col: master.col + master.colSpan - 1,
                },
              });
            }
          }
        }
        rowData.push(row);
      }

      const getCombinedTag = (row) => {
        const epeTag = (row[EPE_COL_INDEX] || "").toString().trim();
        const merTag = (row[MER_COL_INDEX] || "").toString().trim();

        if (epeTag === "NOT IN EPE" && merTag !== "NOT IN MER") {
          return merTag;
        } else if (merTag === "NOT IN MER" && epeTag !== "NOT IN EPE") {
          return epeTag;
        } else if (epeTag !== "NOT IN EPE" && merTag !== "NOT IN MER") {
          return epeTag || merTag; // Use EPE if available, otherwise MER
        } else {
          return epeTag + merTag; // Both are "NOT IN", combine them
        }
      };

      // Step 3: Sort the row data based on the combined tag
      rowData.sort((a, b) => {
        const tagA = getCombinedTag(a);
        const tagB = getCombinedTag(b);
        return tagA.localeCompare(tagB);
      });

      // Step 3: Group rows by merged cells
      const groupedRows = [];
      let currentGroup = [];
      rowData.forEach((row, index) => {
        const isNotIn =
          row[EPE_COL_INDEX] === "NOT IN EPE" &&
          row[MER_COL_INDEX] === "NOT IN MER";

        if (isNotIn) {
          // If the current row is a "NOT IN" row, add it as a single-row group
          if (currentGroup.length > 0) {
            groupedRows.push(currentGroup);
            currentGroup = [];
          }
          groupedRows.push([{ row, originalIndex: index + 2 }]);
        } else {
          currentGroup.push({ row, originalIndex: index + 2 });
          const nextRow = rowData[index + 1];
          if (
            !nextRow ||
            (row[EPE_COL_INDEX] !== nextRow[EPE_COL_INDEX] &&
              row[MER_COL_INDEX] !== nextRow[MER_COL_INDEX])
          ) {
            groupedRows.push(currentGroup);
            currentGroup = [];
          }
        }
      });

      // Add any remaining rows in the current group
      if (currentGroup.length > 0) {
        groupedRows.push(currentGroup);
      }

      // Step 4: Sort groups by drawing number
      groupedRows.sort((a, b) => {
        const drawingA = a[0].row[DRAWING_NUMBER_COL_INDEX] || "";
        const drawingB = b[0].row[DRAWING_NUMBER_COL_INDEX] || "";
        return drawingA.localeCompare(drawingB);
      });

      // Step 5: Flatten groups and rewrite data to worksheet
      let newRowIndex = 2;
      groupedRows.forEach((group) => {
        group.forEach(({ row, originalIndex }) => {
          for (let j = 1; j <= worksheet.columnCount; j++) {
            setCellValue(newRowIndex, j, row[j]);
          }
          newRowIndex++;
        });
      });

      // Function to safely merge cells
      const safeMergeCells = (startRow, endRow, colIndex) => {
        if (startRow === endRow) return; // No need to merge a single cell
        try {
          worksheet.mergeCells(startRow, colIndex, endRow, colIndex);
        } catch (error) {
          // console.error(`Error merging cells from row ${startRow} to ${endRow}, column ${colIndex}: ${error.message}`);
        }
      };

      // Function to merge cells for a specific column and related columns
      const mergeCellsForColumn = (colIndex) => {
        let currentValue = null;
        let startRow = null;
        let endRow = null;
        let slNo = 1;

        for (let i = 2; i <= worksheet.rowCount; i++) {
          const cellValue = getCellValue(i, colIndex);
          const epeValue = getCellValue(i, EPE_COL_INDEX);

          // For MER column, ignore "NOT IN EPE" in EPE column
          const shouldMerge =
            colIndex === MER_COL_INDEX
              ? cellValue === currentValue &&
                cellValue &&
                cellValue !== "NOT IN MER"
              : cellValue === currentValue &&
                cellValue &&
                cellValue !== "NOT IN EPE" &&
                cellValue !== "NOT IN MER";

          if (shouldMerge) {
            endRow = i;
          } else {
            if (startRow !== null && endRow !== null) {
              if (startRow !== endRow) {
                safeMergeCells(startRow, endRow, colIndex);

                columnsToMerge.forEach((relatedColIndex) => {
                  let shouldMerge = true;

                  // Skip content comparison for the first column (e.g., SL.NO column)
                  if (relatedColIndex !== SL_NO_COL_INDEX) {
                    // Loop through each row in the range and check if the contents are the same
                    for (let row = startRow; row <= endRow; row++) {
                      const currentCellValue =
                        worksheet[
                          XLSX.utils.encode_cell({ r: row, c: relatedColIndex })
                        ]?.v;
                      const firstCellValue =
                        worksheet[
                          XLSX.utils.encode_cell({
                            r: startRow,
                            c: relatedColIndex,
                          })
                        ]?.v;

                      // If any cell in the column has a different value, skip merging for this column
                      if (currentCellValue !== firstCellValue) {
                        shouldMerge = false;
                        break;
                      }
                    }
                  }

                  // Don't merge if the contents are different, but always merge SL.NO column
                  if (shouldMerge || relatedColIndex === SL_NO_COL_INDEX) {
                    // For MER column, don't merge EPE column if it contains "NOT IN EPE"
                    if (
                      !(
                        colIndex === MER_COL_INDEX &&
                        relatedColIndex === EPE_COL_INDEX &&
                        epeValue === "NOT IN EPE"
                      )
                    ) {
                      safeMergeCells(startRow, endRow, relatedColIndex);
                    }
                  }
                });

                // Set SL.NO for the merged group
                setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
              } else {
                // Set SL.NO for single row
                setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
              }
            }

            currentValue = cellValue;
            startRow = i;
            endRow = i;
          }
        }

        // Handle the last group
        if (startRow !== null && endRow !== null) {
          if (startRow !== endRow) {
            safeMergeCells(startRow, endRow, colIndex);

            // Merge related columns for the last range including SL.NO
            columnsToMerge.forEach((relatedColIndex) => {
              let shouldMerge = true;

              // Skip content comparison for the first column (e.g., SL.NO column)
              if (relatedColIndex !== SL_NO_COL_INDEX) {
                // Loop through each row in the range and check if the contents are the same
                for (let row = startRow; row <= endRow; row++) {
                  const currentCellValue =
                    worksheet[
                      XLSX.utils.encode_cell({ r: row, c: relatedColIndex })
                    ]?.v;
                  const firstCellValue =
                    worksheet[
                      XLSX.utils.encode_cell({
                        r: startRow,
                        c: relatedColIndex,
                      })
                    ]?.v;

                  // If any cell in the column has a different value, skip merging for this column
                  if (currentCellValue !== firstCellValue) {
                    shouldMerge = false;
                    break;
                  }
                }
              }

              // Don't merge if the contents are different, but always merge SL.NO column
              if (shouldMerge || relatedColIndex === SL_NO_COL_INDEX) {
                // For MER column, don't merge EPE column if it contains "NOT IN EPE"
                if (
                  !(
                    colIndex === MER_COL_INDEX &&
                    relatedColIndex === EPE_COL_INDEX &&
                    epeValue === "NOT IN EPE"
                  )
                ) {
                  safeMergeCells(startRow, endRow, relatedColIndex);
                }
              }
            });

            // Set SL.NO for the last merged group
            setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
          } else {
            // Set SL.NO for the last single row
            setCellValue(startRow, SL_NO_COL_INDEX, slNo++);
          }
        }
      };

      // Step 5: Merge cells for EPE column
      mergeCellsForColumn(EPE_COL_INDEX);

      // Step 6: Merge cells for MER column
      mergeCellsForColumn(MER_COL_INDEX);

      // Step 7: Update SL.NO
      let currentSlNo = 1;
      for (let i = 2; i <= worksheet.rowCount; i++) {
        const cell = worksheet.getCell(i, SL_NO_COL_INDEX);
        if (cell._mergeCount) {
          i = i + cell._mergeCount - 1;
        } else {
          setCellValue(i, SL_NO_COL_INDEX, currentSlNo++);
        }
      }
    };

    // Apply the sorting and merging logic
    sortAndMergeColumns();

    // Generate and download the Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "merged-excel.xlsx";
    link.click();
  };

  return {
    handleFileChange,
    resetUpload,
    processExcelFiles,
    loading,
    uploads,
    processAndMergeSldExcelFiles,
  };
};

export default useExcelProcessor;
