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



const HomeMintues = () => {
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
        email: value.email,
        surname: value.surname,
        totalScore: value.totalScore,
        scores: value.scores.join(', '),
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
          let email,score, name, surname;
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
            else if (columnName === 'Time in Session') {
              score = parseInt(cellValue) || 0;
            }
          });

          if (email) {
            const key = crypto.createHash('md5').update(`${name}`).digest('hex');
            if (dataMap.has(key)) {
              const entry = dataMap.get(key);
              entry.scores.push(score);
              entry.totalScore += score;
             
            } else {
              dataMap.set(key, {
                name,
                email,
                surname,
                totalScore: score,
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
      { header: 'First Name', key: 'name', width: 30 },
      { header: 'Last Name', key: 'surname', width: 30 },
      { header: 'Email Address', key: 'email', width: 30 },
      { header: 'Time in Session ', key: 'scores', width: 50 },
      { header: 'Total Time in Session ', key: 'totalScore', width: 15 },     
    ];

    worksheet.columns = columns;

    const sortedDataArray = Array.from(dataMap.values()).sort((a, b) => {
      // Sorting by name, if available, otherwise by email
      if (a.name && b.name) {
        if (a.name < b.name) return -1;
        if (a.name > b.name) return 1;
      } 
      return 0;
    });
    sortedDataArray.forEach((value) => {
      const rowData = {
        name: value.name,
        surname: value.surname,
        email: value.email,
        totalScore: value.totalScore,
        scores: value.scores.join(', '),
      };
      worksheet.addRow(rowData);
  
    })
   
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
              <TableHead >First Name</TableHead>
              <TableHead >Last Name</TableHead>
              <TableHead >Email</TableHead>
              <TableHead className="">Time in Session</TableHead>
              <TableHead>Total Time in Session</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
          {combinedData.map((item, index) => (
            <TableRow key={index} className='mb-5'>
        
        <TableCell className="font-medium">{item.name}</TableCell>
        <TableCell className="font-medium">{item.surname}</TableCell>
        <TableCell className="font-medium">{item.email}</TableCell>
        
        <TableCell>{item.scores}</TableCell>
        <TableCell >{item.totalScore}</TableCell>
        
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

export default HomeMintues;
