import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';

const App = () => {
  const [data, setData] = useState([]);
  const [cellStyles, setCellStyles] = useState({});

  const handleFileChange = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const ab = e.target.result;
      const wb = XLSX.read(ab, { type: 'array', cellStyles: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const filteredData = jsonData.filter((row, index) => {
        return (
          index === 1 ||
          index === 2 ||
          (row[1] && (row[1].includes("BasiX") || row[1].includes("FieldService")))
        );
      });

      if (filteredData[0]) filteredData[0].unshift('');
      if (filteredData[1]) filteredData[1].unshift('');

      const newCellStyles = {};

      if (filteredData[2]) {
        const row3 = filteredData[2];
        for (let i = 2; i < row3.length - 1; i++) {
          if (row3[i] !== row3[i + 1]) {
            const cellAddress = XLSX.utils.encode_cell({ r: 2, c: i });
            newCellStyles[cellAddress] = 'blue';
          }
        }
      }

      filteredData.forEach((row, rowIndex) => {
        if (row[1] && (row[1].includes("BasiX") || row[1].includes("FieldService"))) {
          for (let i = 2; i < row.length - 1; i++) {
            if (row[i] !== row[i + 1]) {
              const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: i });
              newCellStyles[cellAddress] = 'blue';
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
    if (cellIndex === 0 || cellIndex === 1) {
      return 'white';
    }
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: cellIndex });
    return cellStyles[cellAddress] || 'white';
  };

  const getRow3DataWithStyles = () => {
    if (data.length <= 2) return [];
    return data[2].map((cell, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 2, c: index });
      return {
        value: cell,
        bgColor: cellStyles[cellAddress] || 'white'
      };
    });
  };

  const getGroupedChartData = () => {
    const row3Data = getRow3DataWithStyles();
    const groupedData = [];
    
    for (let i = 3; i < row3Data.length; i += 7) {
      const group = row3Data.slice(i, i + 7);
      const blueCount = group.filter(cell => cell.bgColor === 'blue').length;
      groupedData.push({
        name: `Cells ${i + 1}-${Math.min(i + 7, row3Data.length)}`,
        value: blueCount,
        totalCells: group.length
      });
    }
    
    return groupedData;
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Dashboard</h1>
      <input
        type="file"
        accept=".xlsm,.xlsx, .xls"
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
      <div style={{ marginTop: '20px', padding: '10px', border: '1px solid #ddd' }}>
        <h2>Row 3 Data:</h2>
        {getRow3DataWithStyles().map((cell, index) => (
          <div key={index} style={{ backgroundColor: cell.bgColor, padding: '5px', display: 'inline-block', margin: '2px' }}>
            {cell.value}
          </div>
        ))}
      </div>
      <div style={{ marginTop: '20px', padding: '10px', border: '1px solid #ddd' }}>
        <h2>Grouped Bar Chart (Blue Cells Count):</h2>
        <div style={{ width: '100%', height: 400 }}>
          <ResponsiveContainer>
            <BarChart data={getGroupedChartData()}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis />
              <Tooltip 
                content={({ active, payload, label }) => {
                  if (active && payload && payload.length) {
                    const data = payload[0].payload;
                    return (
                      <div style={{ backgroundColor: '#fff', padding: '5px', border: '1px solid #999' }}>
                        <p>{`${label}`}</p>
                        <p>{`Blue Cells: ${data.value}`}</p>
                        <p>{`Total Cells: ${data.totalCells}`}</p>
                      </div>
                    );
                  }
                  return null;
                }}
              />
              <Bar dataKey="value" fill="#8884d8" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};

export default App;
