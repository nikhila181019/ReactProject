import React, { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LabelList } from 'recharts';
import { fetchExcelFile, getRow3DataWithStyles, getGroupedChartData } from './promotionService';

const Index = () => {
  const [data, setData] = useState([]);
  const [cellStyles, setCellStyles] = useState({});

  useEffect(() => {
    const loadData = async () => {
      try {
        const { data, cellStyles } = await fetchExcelFile();
        setData(data);
        setCellStyles(cellStyles);
      } catch (error) {
        console.error('Error loading data:', error);
      }
    };
    loadData();
  }, []);

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
        <div style={styles.tooltip}>
          <p style={styles.tooltipLabel}>{`${label}`}</p>
          <p style={styles.tooltipContent}>{`Blue Cells: ${data.value}`}</p>
          <p style={styles.tooltipContent}>{`Total Cells: ${data.totalCells}`}</p>
        </div>
      );
    }
    return null;
  };

  const row3Data = getRow3DataWithStyles(data, cellStyles);
  const chartData = getGroupedChartData(row3Data);

  return (
    <div style={styles.container}>
      <h1>PDC Status</h1>
      <div style={styles.chartContainer}>
        <h2 style={styles.chartTitle}>PDC Promotion</h2>
        <h2 style={styles.chartTitle}>Last 7 Weeks Data</h2>
        <div style={styles.chart}>
          <ResponsiveContainer>
            <BarChart
              data={chartData}
              barSize={60}
              barGap={10}
              barCategoryGap="15%"
              margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
            >
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis 
                domain={[0, 10]}
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
              >
                <LabelList content={<CustomizedLabel />} position="top" />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div style={styles.legend}>
          <div style={styles.legendItem}>
            <span style={{...styles.legendDot, backgroundColor: 'red'}}></span>
            <span>PDC Promotion = 0</span>
          </div>
          <div style={styles.legendItem}>
            <span style={{...styles.legendDot, backgroundColor: 'green'}}></span>
            <span>PDC Promotion &gt; 0</span>
          </div>
        </div>
      </div>
    </div>
  );
}

const styles = {
  container: {
    padding: '20px',
    fontFamily: 'Arial, sans-serif',
  },
  chartContainer: {
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
    margin: '0 auto',
  },
  chartTitle: {
    margin: '0 0 20px 0',
    fontSize: '24px',
  },
  chart: {
    width: '100%',
    height: '90%',
  },
  legend: {
    display: 'flex',
    justifyContent: 'center',
    marginTop: '20px',
  },
  legendItem: {
    display: 'flex',
    alignItems: 'center',
    marginRight: '20px',
  },
  legendDot: {
    borderRadius: '50%',
    width: '12px',
    height: '12px',
    display: 'inline-block',
    marginRight: '5px',
  },
  tooltip: {
    backgroundColor: '#fff',
    padding: '10px',
    border: '1px solid #999',
    borderRadius: '4px',
  },
  tooltipLabel: {
    margin: '0',
    fontWeight: 'bold',
  },
  tooltipContent: {
    margin: '5px 0 0',
  },
  table: {
    borderCollapse: 'collapse',
    width: '100%',
  },
  th: {
    border: '1px solid #ddd',
    padding: '8px',
    backgroundColor: '#f2f2f2',
    textAlign: 'left',
  },
  td: {
    border: '1px solid #ddd',
    padding: '8px',
  },
};

export default Index;
