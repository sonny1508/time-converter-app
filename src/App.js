import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function App() {
  const [sourceFile, setSourceFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [logs, setLogs] = useState([]);

  const handleSourceFileChange = (e) => {
    const file = e.target.files[0];
    setSourceFile(file);
  };

  const handleTemplateFileChange = (e) => {
    const file = e.target.files[0];
    setTemplateFile(file);
  };

  const addLog = (message) => {
    setLogs(prevLogs => [...prevLogs, message]);
  };

  const convertTimeFormat = async (sourceData, templateData) => {
    try {
      addLog("Reading source file...");
      // Read the source file
      const sourceWorkbook = XLSX.read(sourceData, {
        cellDates: true,
      });
      const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
      const sourceJson = XLSX.utils.sheet_to_json(sourceSheet);
      addLog(`Found ${sourceJson.length} entries in the source file`);

      addLog("Reading template file...");
      // Read the template file
      const templateWorkbook = XLSX.read(templateData, {
        cellDates: true,
      });
      const templateSheet = templateWorkbook.Sheets[templateWorkbook.SheetNames[0]];
      const templateRange = XLSX.utils.decode_range(templateSheet['!ref']);
      
      // Get the header row to find date columns
      const headerRow = [];
      for (let c = templateRange.s.c; c <= templateRange.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({r: 0, c: c});
        const cell = templateSheet[cellAddress];
        headerRow.push(cell ? cell.v : null);
      }
      
      // Extract employees from template
      const employees = [];
      for (let r = 1; r <= templateRange.e.r; r++) {
        const nameCell = templateSheet[XLSX.utils.encode_cell({r, c: 1})];
        if (nameCell && nameCell.v) {
          employees.push({
            row: r,
            name: nameCell.v
          });
        }
      }
      addLog(`Found ${employees.length} employees in the template`);
      
      // Find date columns in the header row
      const dateColumns = [];
      for (let i = 0; i < headerRow.length; i++) {
        const header = headerRow[i];
        if (header && typeof header === 'string' && header.includes('/')) {
          // For each date, we need the column for login and the next column for logout
          dateColumns.push({
            date: header,
            loginCol: i,
            logoutCol: i + 1
          });
        }
      }
      addLog(`Found ${dateColumns.length} date columns in the template`);
      
      // Group data by employee, date, and session
      addLog("Processing login/logout data...");
      // This ensures we correctly match login with its corresponding logout
      const groupedData = {};
      
      // First, organize data by employee and date
      sourceJson.forEach(entry => {
        const name = entry['Họ tên'];
        if (!name) return;
        
        if (!groupedData[name]) {
          groupedData[name] = {};
        }
        
        if (entry['Thời điểm đăng nhập']) {
          const loginDate = new Date(entry['Thời điểm đăng nhập']);
          const dateKey = `${loginDate.getDate()}/${loginDate.getMonth() + 1}`;
          
          if (!groupedData[name][dateKey]) {
            groupedData[name][dateKey] = [];
          }
          
          // Create a new session entry
          const session = {
            login: loginDate,
            logout: entry['Thời điểm đăng xuất'] ? new Date(entry['Thời điểm đăng xuất']) : null
          };
          
          groupedData[name][dateKey].push(session);
        } else if (entry['Thời điểm đăng xuất']) {
          // Handle case where we only have logout without login (shouldn't happen, but just in case)
          const logoutDate = new Date(entry['Thời điểm đăng xuất']);
          const dateKey = `${logoutDate.getDate()}/${logoutDate.getMonth() + 1}`;
          
          if (!groupedData[name][dateKey]) {
            groupedData[name][dateKey] = [];
          }
          
          // Add session with only logout
          groupedData[name][dateKey].push({
            login: null,
            logout: logoutDate
          });
        }
      });
      
      // Process grouped data to fill template
      addLog("Filling template with processed data...");
      let filledCells = 0;
      
      employees.forEach(employee => {
        const employeeData = groupedData[employee.name];
        
        if (employeeData) {
          dateColumns.forEach(dateCol => {
            const dateKey = dateCol.date;
            
            if (employeeData[dateKey] && employeeData[dateKey].length > 0) {
              // Get the first session for this date (most relevant)
              const session = employeeData[dateKey][0];
              
              // Set login time if available
              if (session.login) {
                const loginTime = `${session.login.getHours().toString().padStart(2, '0')}:${session.login.getMinutes().toString().padStart(2, '0')}`;
                const loginCell = XLSX.utils.encode_cell({r: employee.row, c: dateCol.loginCol});
                templateSheet[loginCell] = {t: 's', v: loginTime};
                filledCells++;
              }
              
              // Set logout time only if it exists and matches the same date
              if (session.logout) {
                const logoutDate = session.logout;
                const sessionDateKey = `${logoutDate.getDate()}/${logoutDate.getMonth() + 1}`;
                
                // Only add logout if it's from the same date as the login
                if (sessionDateKey === dateKey) {
                  const logoutTime = `${logoutDate.getHours().toString().padStart(2, '0')}:${logoutDate.getMinutes().toString().padStart(2, '0')}`;
                  const logoutCell = XLSX.utils.encode_cell({r: employee.row, c: dateCol.logoutCol});
                  templateSheet[logoutCell] = {t: 's', v: logoutTime};
                  filledCells++;
                }
              }
            }
          });
        }
      });
      
      addLog(`Filled ${filledCells} cells with time data`);
      
      // Generate the output workbook
      const outputWorkbook = templateWorkbook;
      addLog("Conversion completed successfully!");
      
      return outputWorkbook;
    } catch (error) {
      console.error('Error converting time format:', error);
      addLog(`Error: ${error.message}`);
      throw error;
    }
  };

  const handleProcess = async () => {
    if (!sourceFile || !templateFile) {
      setError('Please select both source and template files');
      return;
    }

    setProcessing(true);
    setError(null);
    setResult(null);
    setLogs([]);

    try {
      const sourceReader = new FileReader();
      sourceReader.onload = async (e) => {
        const sourceData = e.target.result;
        
        const templateReader = new FileReader();
        templateReader.onload = async (e2) => {
          const templateData = e2.target.result;
          
          try {
            const outputWorkbook = await convertTimeFormat(sourceData, templateData);
            
            // Convert workbook to blob for download
            const outputBlob = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([outputBlob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            // Create download URL
            const url = URL.createObjectURL(blob);
            
            setResult({
              url,
              filename: 'Employee_Time_Converted.xlsx'
            });
            setProcessing(false);
          } catch (err) {
            setError(`Error processing files: ${err.message}`);
            setProcessing(false);
          }
        };
        
        templateReader.readAsArrayBuffer(templateFile);
      };
      
      sourceReader.readAsArrayBuffer(sourceFile);
    } catch (err) {
      setError(`Error reading files: ${err.message}`);
      setProcessing(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto p-6 bg-white rounded-lg shadow-md">
      <h1 className="text-2xl font-bold mb-6 text-center">Employee Time Format Converter</h1>
      
      <div className="space-y-6">
        <div className="bg-blue-50 p-4 rounded-md">
          <p className="text-sm text-blue-800">
            This application converts employee login/logout time data from the NVDangNhap format to the reference format.
            It extracts just the time portion (without date) and places it in the corresponding cells.
            <br /><br />
            <strong>Note:</strong> Only properly paired login/logout times from the same date will be used. 
            Missing logout times will be left blank.
          </p>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="space-y-3">
            <label className="block text-sm font-medium text-gray-700">
              Source File (NVDangNhap format)
            </label>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleSourceFileChange}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
            {sourceFile && (
              <p className="text-xs text-gray-500">Selected: {sourceFile.name}</p>
            )}
          </div>
          
          <div className="space-y-3">
            <label className="block text-sm font-medium text-gray-700">
              Template File (Reference format)
            </label>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleTemplateFileChange}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
            {templateFile && (
              <p className="text-xs text-gray-500">Selected: {templateFile.name}</p>
            )}
          </div>
        </div>
        
        <div className="flex justify-center">
          <button
            onClick={handleProcess}
            disabled={processing || !sourceFile || !templateFile}
            className="px-6 py-2 bg-blue-600 text-white font-medium rounded-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {processing ? 'Processing...' : 'Convert Format'}
          </button>
        </div>
        
        {logs.length > 0 && (
          <div className="mt-4 p-3 bg-gray-50 rounded-md text-sm text-gray-700 max-h-40 overflow-y-auto">
            <h3 className="font-medium mb-2">Processing Log:</h3>
            {logs.map((log, index) => (
              <div key={index} className="mb-1">• {log}</div>
            ))}
          </div>
        )}
        
        {error && (
          <div className="p-4 bg-red-50 text-red-800 rounded-md">
            {error}
          </div>
        )}
        
        {result && (
          <div className="p-4 bg-green-50 text-green-800 rounded-md text-center">
            <p className="mb-3">Conversion completed successfully!</p>
            <a
              href={result.url}
              download={result.filename}
              className="px-4 py-2 bg-green-600 text-white font-medium rounded-md hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2"
            >
              Download Converted File
            </a>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;