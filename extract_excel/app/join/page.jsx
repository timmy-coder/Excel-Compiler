// pages/index.js
'use client'
// pages/index.js
import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import crypto from 'crypto';
import {
  Table,
  TableBody,
  TableCaption,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "../../components/ui/table"



const Home = () => {
  const [files, setFiles] = useState([]);
  const [combinedData, setCombinedData] = useState();
  const [data, setData] = useState()

  const handleFileUpload = (e) => {
    setFiles(e.target.files);
  };

  const handleDownload = () => {
    generateExcel(data);
  }
  const handleProcessFiles = async () => {
    const dataMap = new Map();

    for (const file of files) {
      await processFile(file, dataMap);
    }
    
    handleTable(dataMap)
    setData(dataMap)
    
  };

  const processFile = async (file, dataMap) => {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
  
    workbook.eachSheet((worksheet) => {
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // Skip header row
          let name, email, score, surname;
          row.eachCell((cell) => {
            const columnLetter = cell.address.match(/[A-Z]+/)[0];
            const columnName = formatCellValue(worksheet.getRow(2).getCell(columnLetter).value);
            const cellValue = formatCellValue(cell.value);
  
            if (columnName === 'First Name') {
              name = cellValue;
            } else if (columnName === 'Email Address') {
              email = cellValue.toLowerCase();
            } 
            else if (columnName === 'Surname') {
              surname = cellValue;
            }
             else if (columnName === 'Score') {
              score = parseFloat(cellValue) || 0;
            }
          });
  
          if (email) {
            const key = crypto.createHash('md5').update(`${email}`).digest('hex');
            if (dataMap.has(key)) {
              const entry = dataMap.get(key);
              entry.scores.push(score);
            } else {
              dataMap.set(key, {
                name: name || '',
                surname: surname || '',
                email,
                scores: score !== undefined ? [score] : [],
              });
            }
          }
        }
      });
    });
  };
  
  const handleTable = (dataMap) => {
    let maxScores = 0;
    const combinedArray = [];
    dataMap.forEach((value) => {
      const rowData = {
        name: value.name,
        surname: value.surname,
        email: value.email,
        scores: value.scores, // Store scores as an array
      };
      combinedArray.push(rowData);
    });
    setCombinedData({ data: combinedArray, maxScores });
  };
  const formatCellValue = (value) => {
    if (value instanceof Date) {
      return value.toISOString();
    }
    return value;
  };
  const generateExcel = async (dataMap) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Combined Data');
  
    // Define the columns dynamically based on the number of scores
    const columns = [
      { header: 'Name', key: 'name', width: 30 },
      { header: 'Surname', key: 'surname', width: 15 },
      { header: 'Email', key: 'email', width: 30 },
    ];
  
    // Determine the maximum number of scores to create dynamic columns

    let maxScores = 0;
    dataMap.forEach((value) => {
      if (value.scores.length > maxScores) {
        maxScores = value.scores.length;
      }
    });
  
    // Add columns for each score
    for (let i = 0; i < maxScores; i++) {
      columns.push({ header: `Score ${i + 1}`, key: `score${i + 1}`, width: 15 });
    }
  
    worksheet.columns = columns;

   
  
    // Add rows to the worksheet
    dataMap.forEach((value) => {
      const rowData = {
        name: value.name,
        surname: value.surname,
        email: value.email,
      };
  
      // Add scores to the respective columns
      value.scores.forEach((score, index) => {
        rowData[`score${index + 1}`] = score;
      });
  
      worksheet.addRow(rowData);
    });
  
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer], { type: 'application/octet-stream' }), 'combined_data.xlsx');
  };

  return (
    <div className='mt-5 mx-20'>
    <h1 className='text-center font-bold text-4xl'>Excel Data Combiner</h1>

    <div className='flex items-center justify-center gap-10 my-20'>
      <input className="flex h-9 w-[300px] rounded-md border border-input bg-transparent px-3 py-1 text-sm shadow-sm transition-colors file:border-0 file:bg-transparent file:text-sm file:font-medium placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring disabled:cursor-not-allowed disabled:opacity-50" type="file" multiple accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button className='p-3 rounded-md bg-slate-400 text-white' onClick={handleProcessFiles}>Generate Excel</button>
      {data && (
        <button className='p-3 rounded-md bg-blue-950 text-white' onClick={handleDownload}>Download</button>
      )}
    </div>

    {combinedData && combinedData.data && (
      <Table>
        <TableCaption>Merge all your Excel files.</TableCaption>
        <TableHeader>
          <TableRow>
            <TableHead>Name</TableHead>
            <TableHead>Surname</TableHead>
            <TableHead>Email</TableHead>
            {[...Array(combinedData.maxScores)].map((_, i) => (
              <TableHead key={i}>{`Score ${i + 1}`}</TableHead>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {combinedData.data.map((item, index) => (
            <TableRow key={index} className='mb-5'>
              <TableCell className="font-medium">{item.name}</TableCell>
              <TableCell className="font-medium">{item.surname}</TableCell>
              <TableCell>{item.email}</TableCell>
              {item.scores.map((score, i) => (
                <TableCell key={i}>{score}</TableCell>
              ))}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    )}
  </div>
);
};

export default Home;
