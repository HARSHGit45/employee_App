var dept = '';

function uploadExcel() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert("No file selected!");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Show the modal for user input
        $('#dataModal').modal('show');

        // Event listener for the submit button in the modal
        document.getElementById('submitData').onclick = function() {
            const sheetName = document.getElementById('sheetName').value;
            const daysInMonth = parseInt(document.getElementById('daysInMonth').value);
            const department = document.getElementById('department').value;

            if (!sheetName || isNaN(daysInMonth) || !department) {
                alert("Please fill out all fields.");
                return;
            }

            // Check if the specified sheet exists
            if (!workbook.Sheets[sheetName]) {
                alert(`Sheet ${sheetName} does not exist in the file.`);
                return;
            }

            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Process JSON to match your required structure
            const employeeData = [];

            for (let i = 0; i < json.length; i++) {
                const row = json[i];
                const employeeCodeName = row[0];

                if (employeeCodeName && employeeCodeName.includes("Emp")) {
                    const employeeInfo = employeeCodeName.split(" ");

                    if (employeeInfo.length >= 3) {
                        const empCode = employeeInfo[2];
                        const empName = employeeInfo[3];
                        const daysInfo = [];

                        // Collecting attendance data for the specified number of days
                        for (let day = 1; day <= daysInMonth; day++) {
                            const dayInfo = {
                                "Day": day,
                                "WeekDay": json[i + 1][day],  // Adjust this based on your actual data structure
                                "Status": json[i + 3][day],    // Adjust index based on the actual row for Status
                                "In": json[i + 4][day],        // Adjust index based on actual row for In time
                                "Out": json[i + 7][day]        // Adjust index based on actual row for Out time
                            };
                            daysInfo.push(dayInfo);
                        }

                        const employeeDocument = {
                            "EmployeeCode": empCode,
                            "EmployeeName": empName,
                            "Attendance": daysInfo
                        };

                        employeeData.push(employeeDocument);
                    }
                }
            }

            // Directly load data for further operations instead of downloading it
            attendanceData = employeeData; // Assuming attendanceData is defined globally
            alert("Excel data loaded successfully and ready for further processing!");
            document.getElementById('processingButtons').classList.remove('d-none');
            $('#dataModal').modal('hide'); 
        };
    };

    reader.readAsArrayBuffer(file);
}



// Function to upload JSON file
function uploadJSON() {
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.json';

  fileInput.onchange = event => {
      const file = event.target.files[0];
      if (!file) {
          alert("No file selected!");
          return;
      }

      const reader = new FileReader();
      reader.onload = e => {
          try {
              // Parse the JSON file
              attendanceData = JSON.parse(e.target.result);
              alert("JSON file uploaded successfully!");
             
          } catch (error) {
              alert("Error parsing JSON file: " + error.message);
          }
      };

      reader.readAsText(file);
  };

  fileInput.click(); // Trigger the file input dialog
}



function calculateTotalWorkingDays() {
    if (!attendanceData.length) {
        alert("No JSON file loaded! Please upload a file first.");
        return;
    }

    const workingDaysSummary = attendanceData.map(employee => {
        const empCode = employee.EmployeeCode;
        const empName = employee.EmployeeName;
        const totalWorkingDays = employee.Attendance.reduce((total, record) => {
            if (record.Status.trim().toUpperCase() === "P") return total + 1; // Full present day
            if (record.Status.trim().toUpperCase() === "1/2P 1/2CL") return total + 0.5; // Half present
            return total; // If absent, do not count
        }, 0);

        return { empCode, empName, totalWorkingDays };
    });

    // Sort the summary in descending order based on totalWorkingDays
    workingDaysSummary.sort((a, b) => b.totalWorkingDays - a.totalWorkingDays);

    updateTable(workingDaysSummary, ["Employee Code", "Employee Name", "Total Working Days"]);
}


function calculateSundaysWorked() {
  if (!attendanceData.length) {
      alert("No JSON file loaded! Please upload a file first.");
      return;
  }

  const sundaysSummary = attendanceData.map(employee => {
      const empCode = employee.EmployeeCode;
      const empName = employee.EmployeeName;
      const sundaysWorked = employee.Attendance.reduce((total, record) => {
          return total + (record.WeekDay.trim() === "Sun" && record.Status.trim() === "P" ? 1 : 0);
      }, 0);

      // Only include employees who worked on Sundays
      return sundaysWorked > 0 ? { empCode, empName, sundaysWorked } : null;
  }).filter(employee => employee !== null); // Filter out null values

  // Sort in descending order based on sundaysWorked
  sundaysSummary.sort((a, b) => b.sundaysWorked - a.sundaysWorked);

  updateTable(sundaysSummary, ["Employee Code", "Employee Name", "Sundays Worked"]);
}

// Function to calculate Saturdays worked
function calculateSaturdaysWorked() {
  if (!attendanceData.length) {
      alert("No JSON file loaded! Please upload a file first.");
      return;
  }

  const saturdaysSummary = attendanceData.map(employee => {
      const empCode = employee.EmployeeCode;
      const empName = employee.EmployeeName;
      const saturdaysWorked = employee.Attendance.reduce((total, record) => {
          return total + (record.WeekDay.trim() === "Sat" && record.Status.trim() === "P" ? 1 : 0);
      }, 0);

      // Only include employees who worked on Saturdays
      return saturdaysWorked > 0 ? { empCode, empName, saturdaysWorked } : null;
  }).filter(employee => employee !== null); // Filter out null values

  // Sort in descending order based on saturdaysWorked
  saturdaysSummary.sort((a, b) => b.saturdaysWorked - a.saturdaysWorked);

  updateTable(saturdaysSummary, ["Employee Code", "Employee Name", "Saturdays Worked"]);
}


function calculateAverageHours() {
  if (!attendanceData || attendanceData.length === 0) {
      alert("No attendance data loaded!");
      return;
  }

  const avgHoursSummary = attendanceData.map(employee => {
      const empCode = employee.EmployeeCode;
      const empName = employee.EmployeeName;

      let totalHours = 0;
      let workingDays = 0;

      employee.Attendance.forEach(record => {
          if (record.Status.trim() === "P") {
              const inTimeDecimal = convertTimeToDecimal(record.In);
              const outTimeDecimal = convertTimeToDecimal(record.Out);

              if (inTimeDecimal && outTimeDecimal) {
                  totalHours += (outTimeDecimal - inTimeDecimal);
                  workingDays++;
              }
          }
      });

      const avgHours = workingDays > 0 ? (totalHours / workingDays) : 0;

      let finalAvgHours;

      // Apply custom rounding logic based on decimal part
      if (avgHours % 1 >= 0.55) {
          finalAvgHours = Math.ceil(avgHours); // Round up to the nearest whole number
      } else {
          finalAvgHours = avgHours.toFixed(2); // Keep two decimal places
      }

      return { empCode, empName, avgHours: finalAvgHours };
  });

  // Update the table with average hours summary
  updateTable(avgHoursSummary,["Employee Code", "Employee Name", "Average Hours"]);
}



// Function to calculate total working hours
function calculateTotalWorkingHours() {
  if (!attendanceData.length) {
      alert("No JSON file loaded! Please upload a file first.");
      return;
  }

  const totalHoursSummary = attendanceData.map(employee => {
      const empCode = employee.EmployeeCode;
      const empName = employee.EmployeeName;
      let totalHours = 0;

      employee.Attendance.forEach(record => {
          if (record.In && record.Out && record.Status.trim() === "P") {
              const inTimeDecimal = convertTimeToDecimal(record.In);
              const outTimeDecimal = convertTimeToDecimal(record.Out);

              if (inTimeDecimal && outTimeDecimal) {
                  totalHours += (outTimeDecimal - inTimeDecimal);
              }
          }
      });

      return { empCode, empName, totalHours: totalHours.toFixed(2) };
  });

  updateTable(totalHoursSummary, ["Employee Code", "Employee Name", "Total Hours"]);
}

// Function to calculate average department working hours
function calculateDeptWorkingHours() {
    if (!attendanceData.length) {
        alert("No JSON file loaded! Please upload a file first.");
        return;
    }

    const totalHoursPerEmployee = []; // To store total hours for each employee
    let totalWorkingHours = 0; // To accumulate total hours worked by all employees
    let totalEmployees = 0; // To count the number of employees

    attendanceData.forEach(employee => {
        let employeeTotalHours = 0; // Total hours for the current employee
        let workingDays = 0;

        employee.Attendance.forEach(record => {
            if (record.In && record.Out && record.Status.trim() === "P") {
                const inTime = record.In; // In time in "HH:MM" format
                const outTime = record.Out; // Out time in "HH:MM" format

                // Convert times to decimal hours
                const inHours = convertTimeToDecimal(inTime);
                const outHours = convertTimeToDecimal(outTime);
                
                // Calculate hours worked
                const hoursWorked = outHours - inHours;
                
                if (hoursWorked > 0) { // Only consider positive hours worked
                    employeeTotalHours += hoursWorked; // Add to employee's total hours
                    totalWorkingHours += hoursWorked; // Add to the total working hours
                    workingDays++;
                }
            }
        });

        if (workingDays > 0) {
            totalHoursPerEmployee.push(employeeTotalHours); // Store total hours for this employee
            totalEmployees++; // Increment the employee count
        }
    });

    // Calculate average hours worked
    const averageHours = totalEmployees > 0 ? (totalWorkingHours / totalEmployees).toFixed(2) : 0;

    // Log for debugging
    console.log("Total Working Hours: ", totalWorkingHours);
    console.log("Total Employees: ", totalEmployees);
    console.log("Average Hours: ", averageHours);

    // Update the department summary
    const departmentSummary = [{
        department: dept,
        avgHours: averageHours
    }];

    updateTable(departmentSummary, ["Department", "Average Hours"]);
}


// Function to update the table with dynamic data
function updateTable(data, headers) {
  const tbody = document.querySelector("#outputTable tbody");
  tbody.innerHTML = ""; // Clear existing rows

  // Create header row if headers provided
  if (headers) {
      const headerRow = document.createElement("tr");
      headers.forEach(header => {
          const th = document.createElement("th");
          th.innerText = header;
          headerRow.appendChild(th);
      });
      tbody.appendChild(headerRow);
  }

  // Create rows for data
  data.forEach(row => {
      const tr = document.createElement("tr");
      for (const key in row) {
          const td = document.createElement("td");
          td.innerText = row[key];
          tr.appendChild(td);
      }
      tbody.appendChild(tr);
  });
}



function convertTimeToDecimal(timeString) {
  if (!timeString) return 0; // Handle empty or null strings

  const [hours, minutes] = timeString.split(':').map(Number);
  return hours + (minutes / 60);
}



