import React, { useCallback, useRef, useState } from "react";
import { read, utils, writeFileXLSX } from 'xlsx';
import { TextField, FormControlLabel, Checkbox, Button, Box, Typography, Grid } from '@mui/material';

const App = () => {
  const [__html, setHtml] = useState("");
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [selectedColumns, setSelectedColumns] = useState([]);
  const [isTableReady, setIsTableReady] = useState(false);
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
    let newColumnIndex = updatedData[0].indexOf('File Name');

    if (newColumnIndex === -1) {
      // Add File Name header if it doesn't exist
      updatedData[0].push('File Name');
      newColumnIndex = updatedData[0].length - 1;
    } else {
      // Clear the "File Name" if it already exists
      for (let i = 1; i < updatedData.length; i++) {
        updatedData[i][newColumnIndex] = '';
      }
    }

    if (selectedColumns.length > 0) {
      selectedColumns.forEach(col => {
        // Find the column index
        const colIndex = columns.indexOf(col);

        // Update the "File Name" with selected column values
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
    setIsTableReady(true); // table is ready for export
  }, [data, selectedColumns, columns]);

  const exportFile = useCallback(() => {
    const elt = tbl.current.getElementsByTagName("TABLE")[0];
    const wb = utils.table_to_book(elt);
    writeFileXLSX(wb, "output.xlsx");
  }, [tbl]);

  return (
    <Box sx={{ padding: 4, backgroundColor: '#f5f5f5', borderRadius: 2 }}>
      <Typography variant="h4" gutterBottom>
        Upload and Manage File
      </Typography>
      <TextField
        id="standard-basic"
        variant="standard"
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        sx={{ marginBottom: 2 }}
      />
      {columns.length > 0 && <Box>
        <Typography variant="h6" gutterBottom>
          Select columns in order so that file name generate in same sequence
        </Typography>
        <Grid container spacing={2}>
          {columns.map((column, index) => (
            column !== "File Name" && (
              <Grid item key={index} xs={12} sm={6} md={4}>
                <FormControlLabel
                  control={
                    <Checkbox
                      name={column}
                      value={column}
                      onChange={() => handleCheckboxChange(column)}
                    />
                  }
                  label={column}
                />
              </Grid>
            )
          ))}
        </Grid>
      </Box>}

      <Box sx={{ marginTop: 2 }}>
        {selectedColumns.length > 0 && <Button variant="contained" onClick={handleColumnUpdate} sx={{ marginRight: 2 }}>
          Generate file name column based on selected columns
        </Button>}
        {isTableReady && <Button variant="contained" onClick={exportFile}>
          Export XLSX
        </Button>}
      </Box>
      <Box sx={{ marginTop: 4 }}>
        <div ref={tbl} dangerouslySetInnerHTML={{ __html }} />
      </Box>
    </Box>
  );
}

export default App;