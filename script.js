// HR Leave Management System - Shared JavaScript
// This file contains common functions used by both HR and Employee portals

// Sample employee data structure (fallback if Excel file is not available)
const sampleEmployees = [
    {
        employee_id: "EMP001",
        name: "John Smith",
        department: "Engineering",
        total_annual_leaves: 20,
        sick_leaves: 10,
        leaves_taken: 5,
        leaves_remaining: 15,
        late_entries: ["2023-10-01 09:30", "2023-10-15 10:00", "2023-11-03 09:45"],
        late_deduction_rate: 10.50,
        total_late_deduction: 31.50
    },
    {
        employee_id: "EMP002",
        name: "Sarah Johnson",
        department: "Marketing",
        total_annual_leaves: 22,
        sick_leaves: 12,
        leaves_taken: 7,
        leaves_remaining: 15,
        late_entries: ["2023-09-05 09:45"],
        late_deduction_rate: 10.50,
        total_late_deduction: 10.50
    },
    {
        employee_id: "EMP003",
        name: "Michael Chen",
        department: "Sales",
        total_annual_leaves: 18,
        sick_leaves: 8,
        leaves_taken: 12,
        leaves_remaining: 6,
        late_entries: ["2023-08-10 10:15", "2023-09-22 09:50", "2023-10-05 10:30", "2023-11-10 09:55"],
        late_deduction_rate: 12.00,
        total_late_deduction: 48.00
    }
];

// Function to load SheetJS library dynamically
function loadSheetJS() {
    return new Promise((resolve, reject) => {
        if (window.XLSX) {
            resolve(window.XLSX);
            return;
        }
        
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
        script.onload = () => resolve(window.XLSX);
        script.onerror = reject;
        document.head.appendChild(script);
    });
}

// Function to read Excel file from local folder
async function readExcelFile(filename = 'employee_data.xlsx') {
    try {
        const XLSX = await loadSheetJS();
        
        // Try to read the file
        const response = await fetch(filename);
        if (!response.ok) {
            throw new Error(`File ${filename} not found. Using sample data.`);
        }
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Get first worksheet
        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Process the data
        const employees = jsonData.map(row => {
            // Handle late entries (could be string or array)
            let lateEntries = [];
            if (row.late_entries) {
                if (typeof row.late_entries === 'string') {
                    lateEntries = row.late_entries.split(';').map(entry => entry.trim()).filter(entry => entry);
                } else if (Array.isArray(row.late_entries)) {
                    lateEntries = row.late_entries;
                }
            }
            
            // Calculate leaves remaining if not provided
            const leavesRemaining = row.leaves_remaining !== undefined 
                ? row.leaves_remaining 
                : (row.total_annual_leaves || 0) - (row.leaves_taken || 0);
            
            return {
                employee_id: row.employee_id || row.employeeId || row.id,
                name: row.name || row.employee_name || '',
                department: row.department || '',
                total_annual_leaves: row.total_annual_leaves || row.annual_leaves || 0,
                sick_leaves: row.sick_leaves || 0,
                leaves_taken: row.leaves_taken || row.leaves_used || 0,
                leaves_remaining: leavesRemaining,
                late_entries: lateEntries,
                late_deduction_rate: row.late_deduction_rate || row.deduction_rate || 0,
                total_late_deduction: row.total_late_deduction || row.late_deduction || 0
            };
        });
        
        return employees;
        
    } catch (error) {
        console.warn('Error reading Excel file:', error.message);
        console.log('Using sample employee data instead.');
        return sampleEmployees;
    }
}

// Function to display employee data in a table
function displayEmployeeTable(employees, containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    
    let html = `
        <table class="data-table">
            <thead>
                <tr>
                    <th>Employee ID</th>
                    <th>Name</th>
                    <th>Department</th>
                    <th>Annual Leaves</th>
                    <th>Sick Leaves</th>
                    <th>Leaves Taken</th>
                    <th>Leaves Remaining</th>
                    <th>Late Days</th>
                    <th>Late Deduction</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    employees.forEach(emp => {
        const lateDays = emp.late_entries ? emp.late_entries.length : 0;
        html += `
            <tr>
                <td>${emp.employee_id}</td>
                <td>${emp.name}</td>
                <td>${emp.department}</td>
                <td>${emp.total_annual_leaves}</td>
                <td>${emp.sick_leaves}</td>
                <td>${emp.leaves_taken}</td>
                <td><strong>${emp.leaves_remaining}</strong></td>
                <td>${lateDays}</td>
                <td>$${emp.total_late_deduction.toFixed(2)}</td>
            </tr>
        `;
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Function to display individual employee details
function displayEmployeeDetails(employee, containerId) {
    const container = document.getElementById(containerId);
    if (!container || !employee) return;
    
    const lateDays = employee.late_entries ? employee.late_entries.length : 0;
    
    let lateEntriesHtml = '<ul>';
    if (employee.late_entries && employee.late_entries.length > 0) {
        employee.late_entries.forEach(entry => {
            lateEntriesHtml += `<li>${entry}</li>`;
        });
    } else {
        lateEntriesHtml += '<li>No late entries recorded</li>';
    }
    lateEntriesHtml += '</ul>';
    
    const html = `
        <div class="employee-info">
            <h2>${employee.name} (${employee.employee_id})</h2>
            <p><strong>Department:</strong> ${employee.department}</p>
            
            <div class="leave-summary">
                <div class="summary-card">
                    <h3>Total Annual Leaves</h3>
                    <div class="summary-value">${employee.total_annual_leaves}</div>
                    <div class="summary-unit">days per year</div>
                </div>
                
                <div class="summary-card green">
                    <h3>Sick Leaves Available</h3>
                    <div class="summary-value">${employee.sick_leaves}</div>
                    <div class="summary-unit">days available</div>
                </div>
                
                <div class="summary-card orange">
                    <h3>Leaves Taken</h3>
                    <div class="summary-value">${employee.leaves_taken}</div>
                    <div class="summary-unit">days used</div>
                </div>
                
                <div class="summary-card red">
                    <h3>Leaves Remaining</h3>
                    <div class="summary-value">${employee.leaves_remaining}</div>
                    <div class="summary-unit">days remaining</div>
                </div>
            </div>
            
            <div class="card">
                <h3>Late Attendance Details</h3>
                <p><strong>Total Late Days:</strong> ${lateDays}</p>
                <p><strong>Late Deduction Rate:</strong> $${employee.late_deduction_rate.toFixed(2)} per day</p>
                <p><strong>Total Late Deduction:</strong> $${employee.total_late_deduction.toFixed(2)}</p>
                
                <h4 style="margin-top: 15px;">Late Entry Records:</h4>
                ${lateEntriesHtml}
            </div>
        </div>
    `;
    
    container.innerHTML = html;
}

// Function to populate employee dropdown
function populateEmployeeDropdown(employees, selectId) {
    const select = document.getElementById(selectId);
    if (!select) return;
    
    // Clear existing options
    select.innerHTML = '<option value="">Select Employee</option>';
    
    // Add employee options
    employees.forEach(emp => {
        const option = document.createElement('option');
        option.value = emp.employee_id;
        option.textContent = `${emp.employee_id} - ${emp.name} (${emp.department})`;
        select.appendChild(option);
    });
}

// Function to find employee by ID
function findEmployeeById(employees, employeeId) {
    return employees.find(emp => emp.employee_id === employeeId);
}

// Function to save data to localStorage
function saveToLocalStorage(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify(data));
        return true;
    } catch (error) {
        console.error('Error saving to localStorage:', error);
        return false;
    }
}

// Function to load data from localStorage
function loadFromLocalStorage(key) {
    try {
        const data = localStorage.getItem(key);
        return data ? JSON.parse(data) : null;
    } catch (error) {
        console.error('Error loading from localStorage:', error);
        return null;
    }
}

// Export functions for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        readExcelFile,
        displayEmployeeTable,
        displayEmployeeDetails,
        populateEmployeeDropdown,
        findEmployeeById,
        saveToLocalStorage,
        loadFromLocalStorage,
        sampleEmployees
    };
}
