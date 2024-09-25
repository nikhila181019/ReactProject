import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

const App = () => {
  const [data, setData] = useState([]);
  const [cellStyles, setCellStyles] = useState({});

  const handleFileChange = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const ab = e.target.result;
      const wb = XLSX.read(ab, { type: 'array', cellStyles: true }); // Enable cellStyles
      const ws = wb.Sheets[wb.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Always include Row 2 (index 1), Row 3 (index 2), Row 4 (index 3)
      // And filter other rows where "SWF" is present in Column A (index 0)
      const filteredData = jsonData.filter((row, index) => {
        return index === 1 || index === 2 || index === 3 || (row[0] && row[0].includes("SWF"));
      });

      const newCellStyles = {};

      // Iterate over the filtered rows
      filteredData.forEach((row, rowIndex) => {
        // Only apply the comparison logic for rows that contain "SWF"
        if (row[0] && row[0].includes("SWF")) {
          for (let i = 2; i < row.length - 1; i++) { // Start comparison from column 3 (index 2)
            if (row[i] !== row[i + 1]) {
              const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: i });
              newCellStyles[cellAddress] = 'blue'; // Mark the cell with blue background
            }
          }
        }
      });

      setData(filteredData);
      setCellStyles(newCellStyles);
    };
    reader.readAsArrayBuffer(file);
  };

  const getCellStyle = (rowIndex, cellIndex) => {
    // Skip applying color to column 1 (index 0) and column 2 (index 1)
    if (cellIndex === 0 || cellIndex === 1) {
      return 'white'; // Default color for these columns
    }
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: cellIndex }); // +1 for header row offset
    return cellStyles[cellAddress] || 'white'; // Default color is white if no style
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Dashboard</h1>
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={(event) => handleFileChange(event.target.files[0])}
      />
      {data.length > 0 && (
        <table style={{ marginTop: '20px', borderCollapse: 'collapse', width: '100%' }}>
          <thead>
            <tr>
              {data[0] && data[0].map((header, index) => (
                <th key={index} style={{ border: '1px solid #ddd', padding: '8px' }}>{header}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.slice(1).map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td
                    key={cellIndex}
                    style={{
                      border: '1px solid #ddd',
                      padding: '8px',
                      backgroundColor: getCellStyle(rowIndex, cellIndex),
                      color: getCellStyle(rowIndex, cellIndex) === 'blue' ? 'white' : 'black'
                    }}
                  >
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
};

export default App;
