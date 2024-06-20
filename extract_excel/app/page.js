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
} from "../components/ui/table"



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

  const handleTable = (dataMap) => {
    const combinedArray = [];
    dataMap.forEach((value) => {
      const rowData = {
        name: value.name,
        surname: value.surname,
        email: value.email,
        totalScore: value.totalScore.toFixed(2),
        scores: value.scores.join(', '),
        averageScore: value.averageScore.toFixed(2),
        counts: value.count
      };
      combinedArray.push(rowData);
    });
    setCombinedData(combinedArray);

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
            const columnName = formatCellValue(worksheet.getRow(1).getCell(columnLetter).value);
            const cellValue = formatCellValue(cell.value);
            if (columnName === 'First Name') {
              name = cellValue;
            } else if (columnName === 'Email Address') {
              email = cellValue.toLowerCase();
            } 
            else if (columnName === 'Surname') {
              surname = cellValue;
            } else if (columnName === 'Score') {
              score = parseFloat(cellValue) || 0;
            }
          });

          if (email && !isNaN(score)) {
            const key = crypto.createHash('md5').update(`${email}`).digest('hex');
            if (dataMap.has(key)) {
              const entry = dataMap.get(key);
              entry.scores.push(score);
              entry.totalScore += score;
              entry.count++;
              entry.averageScore = entry.totalScore / entry.count;
            } else {
              dataMap.set(key, {
                name,
                email,
                surname,
                totalScore: score,
                count: 1,
                averageScore: score,
                scores: [score],
              });
            }
          }
        }
      });
    });
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

    const columns = [
      { header: 'Name', key: 'name', width: 30 },
      { header: 'Surname', key: 'surname', width: 15 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Scores', key: 'scores', width: 50 },
      { header: 'Total Score', key: 'totalScore', width: 15 },
      { header: 'Average Score', key: 'averageScore', width: 15 },
      { header: 'Tests Taken', key: 'counts', width: 15 },
     
    ];

    worksheet.columns = columns;
    dataMap.forEach((value) => {
      const rowData = {
        name: value.name,
        surname: value.surname,
        email: value.email,
        totalScore: value.totalScore.toFixed(2),
        scores: value.scores.join(', '),
        averageScore: value.averageScore.toFixed(2),
        counts: value.count
      };
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
      <button className='p-3  rounded-md bg-slate-400 text-white' onClick={handleProcessFiles}>Generate Excel</button>
      {data&&(
        <button  className='p-3  rounded-md bg-blue-950 text-white' onClick={handleDownload}>Download</button>
      )}
      </div>

      {combinedData&&(
          <Table >
          <TableCaption>Merge alll your Excel files.</TableCaption>
          <TableHeader>
            <TableRow>
              <TableHead >Name</TableHead>
              <TableHead>Surname</TableHead>
              <TableHead>Email</TableHead>
              <TableHead className="w-[250px]">Scores</TableHead>
              <TableHead>Total Score</TableHead>
              <TableHead>Average Score</TableHead>
              <TableHead>Tests Taken</TableHead>
              <TableHead className="text-right">Pass/Fail</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
          {combinedData.map((item, index) => (
            <TableRow key={index} className='mb-5'>
        
        <TableCell className="font-medium">{item.name}</TableCell>
        
        <TableCell className="font-medium">{item.surname}</TableCell>
        
        <TableCell >{item.email}</TableCell>
        
        <TableCell>{item.scores}</TableCell>
        <TableCell >{item.totalScore}</TableCell>
        
        <TableCell >{item.averageScore}</TableCell>
        
        <TableCell className='text-center' >{item.counts}</TableCell>
        <TableCell  className={`text-center text-white ${item.counts==7 && item.averageScore >= 80?'bg-green-500 p-2 rounded-md': 'bg-red-600 p-2 rounded-md'}`}>{item.counts==7 && item.averageScore >= 80?'passed': 'fail'}</TableCell>
            </TableRow>
        
                  ))}
        
            <TableRow>
            </TableRow>
          </TableBody>
        </Table>
      )}
    

    </div>
  );
};

export default Home;
