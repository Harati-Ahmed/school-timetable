// Constants
const API_BASE_URL = 'http://localhost:8000/api';
const SPREADSHEET_ID = ''; // Add your Google Spreadsheet ID here

// Utility Functions
const showMessage = (message, type = 'info') => {
    const modal = new bootstrap.Modal(document.getElementById('messageModal'));
    document.getElementById('modalMessage').textContent = message;
    document.getElementById('modalMessage').className = `alert alert-${type}`;
    modal.show();
};

const handleError = (error) => {
    console.error('Error:', error);
    showMessage(error.message || 'An error occurred', 'danger');
};

const fetchApi = async (endpoint, options = {}) => {
    try {
        const response = await fetch(`${API_BASE_URL}${endpoint}`, {
            ...options,
            headers: {
                'Content-Type': 'application/json',
                ...options.headers
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        handleError(error);
        throw error;
    }
};

// Table Management Functions
const createTable = (data, tableId) => {
    const table = document.createElement('table');
    table.className = 'table table-striped table-bordered';
    
    // Create header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body
    const tbody = document.createElement('tbody');
    data.slice(1).forEach(row => {
        const tr = document.createElement('tr');
        row.forEach((cell, index) => {
            const td = document.createElement('td');
            td.textContent = cell;
            if (isBreakOrLunchColumn(index, tableId)) {
                td.className = 'break-column';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    return table;
};

const isBreakOrLunchColumn = (index, tableId) => {
    if (tableId === 'teachersTable') {
        return index === 6 || index === 10; // Break and Lunch columns for teachers
    } else if (tableId === 'classesTable') {
        return index === 5 || index === 9; // Break and Lunch columns for classes
    }
    return false;
};

// Data Loading Functions
const loadTeachersData = async () => {
    try {
        const response = await fetchApi(`/sheets/${SPREADSHEET_ID}?range=Teachers!A1:N`);
        const table = createTable(response.values, 'teachersTable');
        document.getElementById('teachersTable').innerHTML = '';
        document.getElementById('teachersTable').appendChild(table);
    } catch (error) {
        console.error('Error loading teachers data:', error);
    }
};

const loadClassesData = async () => {
    try {
        const response = await fetchApi(`/sheets/${SPREADSHEET_ID}?range=Classes!A1:M`);
        const table = createTable(response.values, 'classesTable');
        document.getElementById('classesTable').innerHTML = '';
        document.getElementById('classesTable').appendChild(table);
    } catch (error) {
        console.error('Error loading classes data:', error);
    }
};

const loadSummaryData = async () => {
    try {
        const response = await fetchApi(`/sheets/${SPREADSHEET_ID}?range=Summary!A1:Z`);
        const table = createTable(response.values, 'summaryTable');
        document.getElementById('summaryTable').innerHTML = '';
        document.getElementById('summaryTable').appendChild(table);
    } catch (error) {
        console.error('Error loading summary data:', error);
    }
};

// Event Handlers
const setupStructure = async () => {
    try {
        await fetchApi(`/timetable/setup-structure?spreadsheet_id=${SPREADSHEET_ID}`, {
            method: 'POST'
        });
        showMessage('Structure setup completed successfully', 'success');
        loadAllData();
    } catch (error) {
        console.error('Error setting up structure:', error);
    }
};

const deployFromConfig = async () => {
    try {
        await fetchApi(`/timetable/deploy-config?spreadsheet_id=${SPREADSHEET_ID}`, {
            method: 'POST'
        });
        showMessage('Configuration deployed successfully', 'success');
        loadAllData();
    } catch (error) {
        console.error('Error deploying config:', error);
    }
};

const updateSummary = async () => {
    try {
        await fetchApi(`/timetable/update-summary?spreadsheet_id=${SPREADSHEET_ID}`, {
            method: 'POST'
        });
        showMessage('Summary updated successfully', 'success');
        loadSummaryData();
    } catch (error) {
        console.error('Error updating summary:', error);
    }
};

const showConfigSheets = async () => {
    try {
        await fetchApi(`/sheets/${SPREADSHEET_ID}/config/show`, {
            method: 'POST'
        });
        showMessage('Configuration sheets are now visible', 'success');
    } catch (error) {
        console.error('Error showing config sheets:', error);
    }
};

const hideConfigSheets = async () => {
    try {
        await fetchApi(`/sheets/${SPREADSHEET_ID}/config/hide`, {
            method: 'POST'
        });
        showMessage('Configuration sheets are now hidden', 'success');
    } catch (error) {
        console.error('Error hiding config sheets:', error);
    }
};

const clearAllData = async () => {
    if (confirm('Are you sure you want to clear all data? This action cannot be undone.')) {
        try {
            await fetchApi(`/sheets/${SPREADSHEET_ID}/clear`, {
                method: 'POST'
            });
            showMessage('All data has been cleared successfully', 'success');
            loadAllData();
        } catch (error) {
            console.error('Error clearing data:', error);
        }
    }
};

const loadAllData = () => {
    loadTeachersData();
    loadClassesData();
    loadSummaryData();
};

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    // Load initial data
    loadAllData();
    
    // Setup event listeners
    document.getElementById('setupStructure').addEventListener('click', setupStructure);
    document.getElementById('deployFromConfig').addEventListener('click', deployFromConfig);
    document.getElementById('refreshSummary').addEventListener('click', updateSummary);
    document.getElementById('showConfigSheets').addEventListener('click', showConfigSheets);
    document.getElementById('hideConfigSheets').addEventListener('click', hideConfigSheets);
    document.getElementById('clearAllData').addEventListener('click', clearAllData);
    
    // Tab change listeners
    const tabs = document.querySelectorAll('a[data-bs-toggle="tab"]');
    tabs.forEach(tab => {
        tab.addEventListener('shown.bs.tab', (event) => {
            const target = event.target.getAttribute('href');
            switch (target) {
                case '#teachers':
                    loadTeachersData();
                    break;
                case '#classes':
                    loadClassesData();
                    break;
                case '#summary':
                    loadSummaryData();
                    break;
            }
        });
    });
});
