import * as XLSX from 'xlsx';

export const fetchExcelFile = async () => {
  try {
    const response = await fetch('/Azurion3.2_PDCStatusReport 1.xlsx');
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    return processExcelData(data);
  } catch (error) {
    console.error('Error fetching the Excel file:', error);
    throw error;
  }
};

const processExcelData = (data) => {
  const wb = XLSX.read(data, { type: 'array', cellStyles: true });
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

  const cellStyles = calculateCellStyles(filteredData);

  return { data: filteredData, cellStyles };
};

const calculateCellStyles = (filteredData) => {
  const cellStyles = {};

  if (filteredData[2]) {
    const row3 = filteredData[2];
    for (let i = 2; i < row3.length - 1; i++) {
      if (row3[i] !== row3[i + 1]) {
        const cellAddress = XLSX.utils.encode_cell({ r: 2, c: i });
        cellStyles[cellAddress] = 'blue';
      }
    }
  }

  filteredData.forEach((row, rowIndex) => {
    if (row[1] && (row[1].includes("BasiX") || row[1].includes("FieldService"))) {
      for (let i = 2; i < row.length - 1; i++) {
        if (row[i] !== row[i + 1]) {
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: i });
          cellStyles[cellAddress] = 'blue';
        }
      }
    }
  });

  return cellStyles;
};

export const getRow3DataWithStyles = (data, cellStyles) => {
  if (data.length <= 2) return [];
  return data[2].map((cell, index) => {
    const cellAddress = XLSX.utils.encode_cell({ r: 2, c: index });
    return {
      value: cell,
      bgColor: cellStyles[cellAddress] || 'white'
    };
  });
};

export const getGroupedChartData = (row3Data) => {
  const groupedData = [];
  const startingWeek = 83; // Assume Week 83 as the latest

  for (let i = 3, week = startingWeek; i < row3Data.length; i += 7, week--) {
    const group = row3Data.slice(i, i + 7);
    const blueCount = group.filter(cell => cell.bgColor === 'blue').length;
    
    groupedData.push({
      name: `Week ${week}`,
      value: blueCount * 2,
      totalCells: group.length
    });

    if (groupedData.length >= 7) break; // Limit to the first 7 groups (weeks)
  }

  return groupedData;
};
