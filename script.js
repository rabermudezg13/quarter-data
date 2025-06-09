document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const results = document.getElementById('results');
    const resultsContent = document.getElementById('resultsContent');
    const downloadBtn = document.getElementById('downloadBtn');

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    // Highlight drop zone when dragging over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });

    // Handle dropped files
    dropZone.addEventListener('drop', handleDrop, false);
    fileInput.addEventListener('change', handleFileSelect, false);

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function highlight(e) {
        dropZone.classList.add('highlight');
    }

    function unhighlight(e) {
        dropZone.classList.remove('highlight');
    }

    function handleDrop(e) {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }

    function handleFileSelect(e) {
        const files = e.target.files;
        handleFiles(files);
    }

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                file.type === 'application/vnd.ms-excel') {
                fileInfo.textContent = `Selected file: ${file.name}`;
                processExcelFile(file);
            } else {
                fileInfo.textContent = 'Please select a valid Excel file (.xlsx or .xls)';
            }
        }
    }

    function processExcelFile(file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get the first worksheet
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            // Process the data
            const processedData = processData(jsonData);
            
            // Display results
            displayResults(processedData);
        };
        reader.readAsArrayBuffer(file);
    }

    function processData(data) {
        // Create a map to store unique combinations
        const uniqueCombinations = new Map();
        
        // Process each row
        data.forEach(row => {
            const talentId = row['Talent Id'];
            const dateWorked = row['Date Worked'];
            const earningsTotal = row['Earnings Total'];
            
            if (talentId && dateWorked) {
                const key = `${talentId}-${dateWorked}-${earningsTotal}`;
                if (!uniqueCombinations.has(key)) {
                    uniqueCombinations.set(key, row);
                }
            }
        });

        // Convert to array and remove duplicates by Talent Id + Date Worked
        const uniqueByDate = new Map();
        Array.from(uniqueCombinations.values()).forEach(row => {
            const key = `${row['Talent Id']}-${row['Date Worked']}`;
            if (!uniqueByDate.has(key)) {
                uniqueByDate.set(key, row);
            }
        });

        // Count work days per talent
        const workDaysCount = new Map();
        Array.from(uniqueByDate.values()).forEach(row => {
            const talentId = row['Talent Id'];
            const talentName = row['Talent Name Full'];
            const key = `${talentId}-${talentName}`;
            
            if (!workDaysCount.has(key)) {
                workDaysCount.set(key, {
                    'Talent Id': talentId,
                    'Talent Name': talentName,
                    'Days Worked': 0
                });
            }
            workDaysCount.get(key)['Days Worked']++;
        });

        return Array.from(workDaysCount.values());
    }

    function displayResults(processedData) {
        results.style.display = 'block';
        
        // Create formatted text content for display
        let textContent = 'Work Days Report\n';
        textContent += '================\n\n';
        
        processedData.forEach(row => {
            textContent += `Talent ID: ${row['Talent Id']}\n`;
            textContent += `Name: ${row['Talent Name']}\n`;
            textContent += `Days Worked: ${row['Days Worked']}\n`;
            textContent += '----------------\n';
        });

        resultsContent.textContent = textContent;
        downloadBtn.style.display = 'block';
        
        // Set up download functionality for Excel
        downloadBtn.onclick = () => {
            // Create a new workbook
            const wb = XLSX.utils.book_new();
            
            // Convert processed data to worksheet
            const ws = XLSX.utils.json_to_sheet(processedData);
            
            // Add column headers
            const headers = ['Talent Id', 'Talent Name', 'Days Worked'];
            XLSX.utils.sheet_add_aoa(ws, [headers], { origin: 'A1' });
            
            // Add copyright notice
            const copyrightNotice = [
                [''],
                ['© Rodrigo Bermudez 2025'],
                ['For exclusive use of Kelly Education'],
                ['No information is stored on external servers and will not be shared']
            ];
            XLSX.utils.sheet_add_aoa(ws, copyrightNotice, { origin: { r: processedData.length + 2, c: 0 } });
            
            // Set column widths
            const colWidths = [
                { wch: 15 }, // Talent Id
                { wch: 30 }, // Talent Name
                { wch: 12 }  // Days Worked
            ];
            ws['!cols'] = colWidths;
            
            // Add the worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, ws, 'Work Days Report');
            
            // Generate Excel file
            XLSX.writeFile(wb, 'Work_Days_Report.xlsx');
        };
    }
}); 