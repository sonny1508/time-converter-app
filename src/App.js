import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function App() {
  const [sourceFile, setSourceFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);
  const [logs, setLogs] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const handleSourceFileChange = (e) => {
    const file = e.target.files[0];
    setSourceFile(file);
  };

  const handleTemplateFileChange = (e) => {
    const file = e.target.files[0];
    setTemplateFile(file);
  };

  const handleStartDateChange = (e) => {
    setStartDate(e.target.value);
  };

  const handleEndDateChange = (e) => {
    setEndDate(e.target.value);
  };

  const addLog = (message) => {
    setLogs(prevLogs => [...prevLogs, message]);
  };

  // Generate array of dates between start and end
  const getDateRange = (start, end) => {
    const dates = [];
    const currentDate = new Date(start);
    const endDate = new Date(end);
    
    while (currentDate <= endDate) {
      dates.push(new Date(currentDate));
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    return dates;
  };

  const convertTimeFormat = async (sourceData, templateData, dateRange) => {
    try {
      addLog("Reading source file...");
      // Read the source file
      const sourceWorkbook = XLSX.read(sourceData, {
        cellDates: true,
        dateNF: 'mm/dd/yyyy hh:mm:ss'
      });
      const sourceSheet = sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]];
      const sourceJson = XLSX.utils.sheet_to_json(sourceSheet, { raw: false, dateNF: 'mm/dd/yyyy hh:mm' });
      addLog(`Found ${sourceJson.length} entries in the source file`);
      
      // Debug: Check column names in source file
      if (sourceJson.length > 0) {
        const columns = Object.keys(sourceJson[0]);
        addLog(`Source columns found: ${columns.join(', ')}`);
        
        // Show first row as example
        const firstRow = sourceJson[0];
        addLog(`First row example: Name="${firstRow[columns[0]] || firstRow['Họ tên'] || firstRow['Họ Tên']}", Login="${firstRow[columns[1]] || firstRow['Thời điểm đăng nhập']}", Logout="${firstRow[columns[2]] || firstRow['Thời điểm đăng xuất']}"`);
      }

      addLog("Reading template file...");
      // Read the template file
      const templateWorkbook = XLSX.read(templateData, {
        cellDates: true,
      });
      const templateSheet = templateWorkbook.Sheets[templateWorkbook.SheetNames[0]];
      const templateRange = XLSX.utils.decode_range(templateSheet['!ref']);
      
      // Extract employees from template (assuming names are in the first or second column)
      const employees = [];
      let nameColumn = 0; // Try first column
      
      // Check if first column has names, otherwise try second column
      for (let r = 1; r <= templateRange.e.r; r++) {
        let nameCell = templateSheet[XLSX.utils.encode_cell({r, c: 0})];
        if (!nameCell || !nameCell.v || nameCell.v === '' || !isNaN(nameCell.v)) {
          nameColumn = 1; // Try second column if first is empty or just numbers
          break;
        }
      }
      
      for (let r = 1; r <= templateRange.e.r; r++) {
        const nameCell = templateSheet[XLSX.utils.encode_cell({r, c: nameColumn})];
        if (nameCell && nameCell.v && nameCell.v.toString().trim() !== '') {
          const cleanName = nameCell.v.toString().trim();
          employees.push({
            row: r,
            name: cleanName
          });
        }
      }
      addLog(`Found ${employees.length} employees in the template`);
      
      // Debug: Show first few employee names from template
      if (employees.length > 0) {
        const sampleNames = employees.slice(0, 3).map(e => e.name).join(', ');
        addLog(`Template employees sample: ${sampleNames}${employees.length > 3 ? '...' : ''}`);
      }
      
      // Process source data to extract login/logout times and working hours by employee and date
      const employeeTimeData = {};
      
      // Try to find column names with various possible formats
      const findColumn = (entry, possibleNames) => {
        for (const name of possibleNames) {
          if (entry.hasOwnProperty(name)) {
            return entry[name];
          }
        }
        return null;
      };
      
      let processedCount = 0;
      let skippedCount = 0;
      
      sourceJson.forEach((entry, index) => {
        // Try various possible column names for employee name
        const name = findColumn(entry, ['Họ tên', 'Họ Tên', 'Ho ten', 'Name', 'Tên', 'Họ và tên']) || 
                    Object.values(entry)[0]; // Fallback to first column
        
        if (!name || name.toString().trim() === '') {
          skippedCount++;
          return;
        }
        
        const employeeName = name.toString().trim();
        
        if (!employeeTimeData[employeeName]) {
          employeeTimeData[employeeName] = {};
        }
        
        // Helper function to parse date and format the date key
        const parseDateToKey = (dateValue) => {
          try {
            if (!dateValue) return null;
            
            let date;
            if (dateValue instanceof Date) {
              date = dateValue;
            } else if (typeof dateValue === 'string') {
              // Handle string format like "10/21/2025 9:00" or "10/21/2025 18:31"
              const cleanedDate = dateValue.trim();
              date = new Date(cleanedDate);
            } else if (typeof dateValue === 'number') {
              // Handle Excel serial date
              date = new Date((dateValue - 25569) * 86400 * 1000);
            } else {
              return null;
            }
            
            if (!isNaN(date.getTime())) {
              // Format as MM/DD/YYYY for consistency
              const month = (date.getMonth() + 1).toString().padStart(2, '0');
              const day = date.getDate().toString().padStart(2, '0');
              const year = date.getFullYear();
              return `${month}/${day}/${year}`;
            }
          } catch (err) {
            if (index < 3) { // Only log first few errors to avoid spam
              addLog(`Date parsing error: ${err.message} for value: ${dateValue}`);
            }
            return null;
          }
          return null;
        };
        
        // Helper function to extract time from date
        const extractTime = (dateValue) => {
          try {
            if (!dateValue) return null;
            
            let date;
            if (dateValue instanceof Date) {
              date = dateValue;
            } else if (typeof dateValue === 'string') {
              const cleanedDate = dateValue.trim();
              // Check if it's just a time (HH:MM format)
              if (cleanedDate.match(/^\d{1,2}:\d{2}$/)) {
                const [hours, minutes] = cleanedDate.split(':');
                return `${hours.padStart(2, '0')}:${minutes}`;
              }
              date = new Date(cleanedDate);
            } else if (typeof dateValue === 'number') {
              date = new Date((dateValue - 25569) * 86400 * 1000);
            } else {
              return null;
            }
            
            if (!isNaN(date.getTime())) {
              return `${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}`;
            }
          } catch (err) {
            return null;
          }
          return null;
        };
        
        // Try to get login time with various column names
        const loginValue = findColumn(entry, ['Thời điểm đăng nhập', 'Thoi diem dang nhap', 'Login', 'Đăng nhập']) ||
                          Object.values(entry)[1]; // Fallback to second column
        
        if (loginValue) {
          const dateKey = parseDateToKey(loginValue);
          const loginTime = extractTime(loginValue);
          
          if (dateKey && loginTime) {
            if (!employeeTimeData[employeeName][dateKey]) {
              employeeTimeData[employeeName][dateKey] = {};
            }
            employeeTimeData[employeeName][dateKey].login = loginTime;
            if (index === 0) {
              addLog(`First entry parsed - ${employeeName}: ${dateKey} login at ${loginTime}`);
            }
            processedCount++;
          }
        }
        
        // Try to get logout time with various column names
        const logoutValue = findColumn(entry, ['Thời điểm đăng xuất', 'Thoi diem dang xuat', 'Logout', 'Đăng xuất']) ||
                           Object.values(entry)[2]; // Fallback to third column
        
        if (logoutValue) {
          const dateKey = parseDateToKey(logoutValue);
          const logoutTime = extractTime(logoutValue);
          
          if (dateKey && logoutTime) {
            if (!employeeTimeData[employeeName][dateKey]) {
              employeeTimeData[employeeName][dateKey] = {};
            }
            employeeTimeData[employeeName][dateKey].logout = logoutTime;
          }
        }
        
        // Try to get working hours with various column names
        const hoursValue = findColumn(entry, ['Số giờ làm', 'So gio lam', 'Hours', 'Total Hours', 'Giờ làm']) ||
                          Object.values(entry)[3] || // Fallback to fourth column
                          Object.values(entry)[4]; // Sometimes it might be fifth column
        
        if (hoursValue !== undefined && hoursValue !== null && hoursValue !== '') {
          // Get date key from either login or logout time
          let dateKey = null;
          if (loginValue) {
            dateKey = parseDateToKey(loginValue);
          } else if (logoutValue) {
            dateKey = parseDateToKey(logoutValue);
          }
          
          if (dateKey) {
            if (!employeeTimeData[employeeName][dateKey]) {
              employeeTimeData[employeeName][dateKey] = {};
            }
            // Store working hours as is (should be a number like 9.5)
            employeeTimeData[employeeName][dateKey].hours = parseFloat(hoursValue) || hoursValue;
          }
        }
      });
      
      addLog(`Processed ${processedCount} time entries, skipped ${skippedCount} invalid rows`);
      
      // Debug: Show sample of extracted data
      const sampleEmployees = Object.keys(employeeTimeData).slice(0, 3);
      sampleEmployees.forEach(emp => {
        const dates = Object.keys(employeeTimeData[emp]).slice(0, 2);
        if (dates.length > 0) {
          addLog(`Sample - ${emp}: ${dates.map(d => `${d}(${employeeTimeData[emp][d].login || 'no-login'}/${employeeTimeData[emp][d].logout || 'no-logout'}/${employeeTimeData[emp][d].hours || 'no-hours'}h)`).join(', ')}`);
        }
      });
      
      // Create new workbook with dynamic columns based on date range
      const newWorkbook = XLSX.utils.book_new();
      const newSheet = {};
      
      // Copy existing headers from template (employee info columns)
      const existingHeaders = [];
      for (let c = 0; c <= templateRange.e.c; c++) {
        const headerCell = templateSheet[XLSX.utils.encode_cell({r: 0, c: c})];
        if (headerCell && headerCell.v) {
          existingHeaders.push(headerCell.v);
          newSheet[XLSX.utils.encode_cell({r: 0, c: c})] = headerCell;
        }
      }
      
      const baseColumnCount = existingHeaders.length;
      let currentColumn = baseColumnCount;
      
      // Add date columns headers
      dateRange.forEach(date => {
        const dateStr = `${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}/${date.getFullYear()}`;
        
        // Add headers for Login, Logout, Total Hours
        newSheet[XLSX.utils.encode_cell({r: 0, c: currentColumn})] = {t: 's', v: `${dateStr} - Login`};
        newSheet[XLSX.utils.encode_cell({r: 0, c: currentColumn + 1})] = {t: 's', v: `${dateStr} - Logout`};
        newSheet[XLSX.utils.encode_cell({r: 0, c: currentColumn + 2})] = {t: 's', v: `${dateStr} - Total Hours`};
        
        currentColumn += 3;
      });
      
      // Copy employee data and add time data
      let filledCells = 0;
      let filledHoursCells = 0;
      let matchedEmployees = 0;
      let unmatchedEmployees = [];
      
      // Create a normalized name map for better matching
      const normalizedTimeData = {};
      Object.keys(employeeTimeData).forEach(name => {
        const normalized = name.toLowerCase().trim().replace(/\s+/g, ' ');
        normalizedTimeData[normalized] = employeeTimeData[name];
      });
      
      employees.forEach((employee, index) => {
        // Copy existing employee data from template
        for (let c = 0; c < baseColumnCount; c++) {
          const cell = templateSheet[XLSX.utils.encode_cell({r: employee.row, c: c})];
          if (cell) {
            newSheet[XLSX.utils.encode_cell({r: index + 1, c: c})] = cell;
          }
        }
        
        // Try to find matching employee data with flexible name matching
        const employeeName = employee.name.toString().trim();
        const normalizedName = employeeName.toLowerCase().replace(/\s+/g, ' ');
        
        // Try exact match first, then normalized match
        let timeData = employeeTimeData[employeeName] || 
                      employeeTimeData[employee.name] || 
                      normalizedTimeData[normalizedName];
        
        if (timeData) {
          matchedEmployees++;
          let dateColumn = baseColumnCount;
          
          dateRange.forEach(date => {
            const dateKey = `${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}/${date.getFullYear()}`;
            
            if (timeData[dateKey]) {
              // Set login time
              if (timeData[dateKey].login) {
                newSheet[XLSX.utils.encode_cell({r: index + 1, c: dateColumn})] = {t: 's', v: timeData[dateKey].login};
                filledCells++;
              }
              
              // Set logout time
              if (timeData[dateKey].logout) {
                newSheet[XLSX.utils.encode_cell({r: index + 1, c: dateColumn + 1})] = {t: 's', v: timeData[dateKey].logout};
                filledCells++;
              }
              
              // Set working hours
              if (timeData[dateKey].hours !== undefined && timeData[dateKey].hours !== null) {
                newSheet[XLSX.utils.encode_cell({r: index + 1, c: dateColumn + 2})] = {t: 'n', v: parseFloat(timeData[dateKey].hours) || 0};
                filledHoursCells++;
              }
            }
            
            dateColumn += 3;
          });
        } else {
          unmatchedEmployees.push(employeeName);
        }
      });
      
      // Log matching statistics
      addLog(`Matched ${matchedEmployees} out of ${employees.length} employees`);
      if (unmatchedEmployees.length > 0 && unmatchedEmployees.length <= 5) {
        addLog(`Unmatched employees: ${unmatchedEmployees.join(', ')}`);
      } else if (unmatchedEmployees.length > 5) {
        addLog(`Unmatched employees (first 5): ${unmatchedEmployees.slice(0, 5).join(', ')}... and ${unmatchedEmployees.length - 5} more`);
      }
      
      // Set the range for the new sheet
      const maxRow = employees.length;
      const maxCol = currentColumn - 1;
      newSheet['!ref'] = XLSX.utils.encode_range({r: 0, c: 0}, {r: maxRow, c: maxCol});
      
      // Add sheet to workbook
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Employee Time Data');
      
      addLog(`Created ${dateRange.length} date columns (${dateRange.length * 3} total columns)`);
      addLog(`Filled ${filledCells} cells with time data`);
      addLog(`Filled ${filledHoursCells} cells with working hours data`);
      addLog("Conversion completed successfully!");
      
      return newWorkbook;
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

    if (!startDate || !endDate) {
      setError('Please select both start and end dates');
      return;
    }

    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (start > end) {
      setError('Start date must be before or equal to end date');
      return;
    }

    setProcessing(true);
    setError(null);
    setResult(null);
    setLogs([]);

    try {
      const dateRange = getDateRange(startDate, endDate);
      addLog(`Processing data for ${dateRange.length} days: ${startDate} to ${endDate}`);
      
      const sourceReader = new FileReader();
      sourceReader.onload = async (e) => {
        const sourceData = e.target.result;
        
        const templateReader = new FileReader();
        templateReader.onload = async (e2) => {
          const templateData = e2.target.result;
          
          try {
            const outputWorkbook = await convertTimeFormat(sourceData, templateData, dateRange);
            
            // Convert workbook to blob for download
            const outputBlob = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([outputBlob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            // Create download URL
            const url = URL.createObjectURL(blob);
            
            // Generate filename with date range
            const startStr = startDate.replace(/-/g, '');
            const endStr = endDate.replace(/-/g, '');
            const filename = `Employee_Time_${startStr}_to_${endStr}.xlsx`;
            
            setResult({
              url,
              filename
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
            This application converts employee login/logout time data from the NVDangNhap format to a dynamically generated Excel file.
            <br /><br />
            <strong>How it works:</strong>
            <br />• Upload the source file (NVDangNhap format) with login/logout times
            <br />• Upload the reference file with employee names and order
            <br />• Select the date range you want to extract
            <br />• The output will have 3 columns per date: Login, Logout, and Total Hours
            <br /><br />
            <strong>Expected format:</strong> Dates in source file should be like "10/1/2025 9:06:08 AM"
          </p>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="space-y-3">
            <label className="block text-sm font-medium text-gray-700">
              Source File (NVDangNhap format)
            </label>
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={handleSourceFileChange}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
            {sourceFile && (
              <p className="text-xs text-gray-500">Selected: {sourceFile.name}</p>
            )}
          </div>
          
          <div className="space-y-3">
            <label className="block text-sm font-medium text-gray-700">
              Reference File (Employee names/order)
            </label>
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={handleTemplateFileChange}
              className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            />
            {templateFile && (
              <p className="text-xs text-gray-500">Selected: {templateFile.name}</p>
            )}
          </div>
        </div>
        
        <div className="space-y-3">
          <label className="block text-sm font-medium text-gray-700">
            Select Date Range to Extract
          </label>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs text-gray-600 mb-1">Start Date</label>
              <input
                type="date"
                value={startDate}
                onChange={handleStartDateChange}
                className="block w-full p-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
              />
            </div>
            <div>
              <label className="block text-xs text-gray-600 mb-1">End Date</label>
              <input
                type="date"
                value={endDate}
                onChange={handleEndDateChange}
                className="block w-full p-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
              />
            </div>
          </div>
          {startDate && endDate && (
            <p className="text-xs text-gray-600 mt-2">
              Will extract data for {Math.floor((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1} days
            </p>
          )}
        </div>
        
        <div className="flex justify-center">
          <button
            onClick={handleProcess}
            disabled={processing || !sourceFile || !templateFile || !startDate || !endDate}
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
              Download {result.filename}
            </a>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;