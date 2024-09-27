import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LabelList } from 'recharts';

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

      if (groupedData.length >= 7) break; // Limit to the first 7 groups (weeks)
    }

    return groupedData;
  };

  const CustomizedLabel = (props) => {
    const { x, y, width, value } = props;
    return (
      <text
        x={x + width / 2}
        y={y - 10}
        fill={value > 0 ? "green" : "red"}
        textAnchor="middle"
        dominantBaseline="middle"
        fontSize={16}
        fontWeight="bold"
      >
        {value}
      </text>
    );
  };

  const CustomTooltip = ({ active, payload, label }) => {
    if (active && payload && payload.length) {
      const data = payload[0].payload;
      return (
        <div style={{ backgroundColor: '#fff', padding: '10px', border: '1px solid #999', borderRadius: '4px' }}>
          <p style={{ margin: '0', fontWeight: 'bold' }}>{`${label}`}</p>
          <p style={{ margin: '5px 0 0' }}>{`Blue Cells: ${data.value}`}</p>
          <p style={{ margin: '5px 0 0' }}>{`Total Cells: ${data.totalCells}`}</p>
        </div>
      );
    }
    return null;
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel Dashboard</h1>
      <input
        type="file"
        accept=".xlsm,.xlsx, .xls"
        onChange={(event) => handleFileChange(event.target.files[0])}
      />
      <div style={{ 
        marginTop: '20px', 
        padding: '20px', 
        border: '1px solid #ddd', 
        borderRadius: '8px',
        boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
        display: 'flex', 
        flexDirection: 'column', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '600px',
        width: '95%',
        maxWidth: '1200px',
        margin: '0 auto'
      }}>
        <h2 style={{ margin: '0 0 20px 0', fontSize: '24px' }}>PDC Promotion:</h2>
        <div style={{ width: '100%', height: '90%' }}>
          <ResponsiveContainer>
            <BarChart
              data={getGroupedChartData()}
              barSize={60}
              barGap={10}
              barCategoryGap="15%"
              margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
            >
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis 
                domain={[0, 5]} 
                tickCount={6} 
                label={{ value: 'PDC Promoted', angle: -90, position: 'insideLeft', offset: -5 }}
              />
              <Tooltip 
                content={<CustomTooltip />} 
                cursor={false}
              />
              <Bar 
                dataKey="value" 
                fill="#4CAF50"
                onMouseEnter={(data, index) => {
                  // You can add custom behavior here when mouse enters a bar
                }}
                onMouseLeave={(data, index) => {
                  // You can add custom behavior here when mouse leaves a bar
                }}
              >
                <LabelList content={<CustomizedLabel />} position="top" />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
          <div style={{ display: 'flex', alignItems: 'center', marginRight: '20px' }}>
            <span style={{ backgroundColor: 'red', borderRadius: '50%', width: '12px', height: '12px', display: 'inline-block', marginRight: '5px' }}></span>
            <span>PDC Promotion &lt; 0</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <span style={{ backgroundColor: 'green', borderRadius: '50%', width: '12px', height: '12px', display: 'inline-block', marginRight: '5px' }}></span>
            <span>PDC Promotion &gt; 0</span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
