import React, { useCallback, useRef, useState } from "react";
import { read, utils, writeFileXLSX } from 'xlsx';

export default function SheetJSReactHTML() {
  const [__html, setHtml] = useState("");
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const tbl = useRef(null);

  const handleFileUpload = useCallback((event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target.result;
      const wb = read(arrayBuffer); // parse the array buffer
      const ws = wb.Sheets[wb.SheetNames[0]]; // get the first worksheet

      // Convert worksheet to JSON to manipulate data
      const jsonData = utils.sheet_to_json(ws, { header: 1 });
      setData(jsonData); // set data state

      // Extract column headers and set columns state
      const colHeaders = jsonData[0];
      setColumns(colHeaders);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleCheckboxChange = useCallback((col) => {
    setSelectedColumns((prevSelected) => {
      if (prevSelected.includes(col)) {
        return prevSelected.filter(item => item !== col);
      } else {
        return [...prevSelected, col];
      }
    });
  }, []);

  const handleColumnUpdate = useCallback(() => {
    const updatedData = [...data];
    let newColumnIndex = updatedData[0].indexOf('New Column');

    if (newColumnIndex === -1) {
      // Add new column header if it doesn't exist
      updatedData[0].push('New Column');
      newColumnIndex = updatedData[0].length - 1;
    } else {
      // Clear the "New Column" if it already exists
      for (let i = 1; i < updatedData.length; i++) {
        updatedData[i][newColumnIndex] = '';
      }
    }

    if (selectedColumns.length > 0) {
      selectedColumns.forEach(col => {
        // Find the column index
        const colIndex = columns.indexOf(col);

        // Update the "New Column" with selected column values
        for (let i = 1; i < updatedData.length; i++) {
          const cellValue = `${updatedData[i][colIndex]}`;
          if (!updatedData[i][newColumnIndex]) {
            updatedData[i][newColumnIndex] = cellValue;
          } else if (!updatedData[i][newColumnIndex].includes(cellValue)) {
            updatedData[i][newColumnIndex] += ` ${cellValue}`;
          }
        }
      });
    }

    // Convert JSON back to worksheet
    const newWs = utils.json_to_sheet(updatedData, { skipHeader: true });

    // Generate HTML from the updated worksheet
    const newHtml = utils.sheet_to_html(newWs);

    setHtml(newHtml); // update state
    setData(updatedData); // update data state
  }, [data, selectedColumns, columns]);

  const exportFile = useCallback(() => {
    const elt = tbl.current.getElementsByTagName("TABLE")[0];
    const wb = utils.table_to_book(elt);
    writeFileXLSX(wb, "SheetJSReactHTML.xlsx");
  }, [tbl]);

  return (
    <>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <div>
        {columns.map((col, index) => (
          col !== "New Column" && (
            <div key={index}>
              <input
                type="checkbox"
                id={`checkbox-${col}`}
                name={col}
                value={col}
                onChange={() => handleCheckboxChange(col)}
              />
              <label htmlFor={`checkbox-${col}`}>{col}</label>
            </div>
          )
        ))}
      </div>
      <button onClick={handleColumnUpdate}>Update Columns</button>
      <button onClick={exportFile}>Export XLSX</button>
      <div ref={tbl} dangerouslySetInnerHTML={{ __html }} />
    </>
  );
}
