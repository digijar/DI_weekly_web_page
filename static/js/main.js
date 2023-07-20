document.addEventListener('DOMContentLoaded', function () {
    const fileInput = document.getElementById('fileInput');
    const previewArea = document.getElementById('previewArea');
    let currentFileIndex = 0;
    let currentWorkbookIndex = 0;
    let workbooksData = [];

    // Initialize the file and workbook dropdowns and buttons
    const fileDropdown = document.createElement('select');
    fileDropdown.addEventListener('change', function () {
        currentFileIndex = fileDropdown.selectedIndex;
        currentWorkbookIndex = 0; // Reset the workbook index when changing files
        updateDropdowns();
        updatePreview();
    });
    const workbookDropdown = document.createElement('select');
    workbookDropdown.addEventListener('change', function () {
        currentWorkbookIndex = workbookDropdown.selectedIndex;
        updatePreview();
    });
    const previousButton = document.createElement('button');
    previousButton.textContent = 'Previous';
    previousButton.addEventListener('click', function () {
        currentFileIndex = Math.max(0, currentFileIndex - 1);
        currentWorkbookIndex = 0; // Reset the workbook index when changing files
        updateDropdowns();
        updatePreview();
    });
    const nextButton = document.createElement('button');
    nextButton.textContent = 'Next';
    nextButton.addEventListener('click', function () {
        currentFileIndex = Math.min(fileInput.files.length - 1, currentFileIndex + 1);
        currentWorkbookIndex = 0; // Reset the workbook index when changing files
        updateDropdowns();
        updatePreview();
    });

    // Function to update the content of the file and workbook dropdowns
    function updateDropdowns() {
        fileDropdown.innerHTML = '';
        for (const file of fileInput.files) {
            const option = document.createElement('option');
            option.textContent = file.name;
            fileDropdown.appendChild(option);
        }
    
        if (workbooksData.length > 0) {
            const workbookData = workbooksData[currentFileIndex];
            if (workbookData.length > 1) {
                workbookDropdown.innerHTML = '';
                for (const workbook of workbookData) {
                    const option = document.createElement('option');
                    option.textContent = workbook.workbookName;
                    workbookDropdown.appendChild(option);
                }
                workbookDropdown.value = workbookData[currentWorkbookIndex].workbookName;
            } else {
                // If the selected file has only one workbook, reset the workbookDropdown
                workbookDropdown.innerHTML = '';
                workbookDropdown.value = '';
            }
        }
    }
    
    fileInput.addEventListener('change', function () {
        previewArea.innerHTML = ''; // Clear previous content
        workbooksData = []; // Clear previous workbook data
    
        const files = fileInput.files;
        if (files.length > 0) {
            for (let i = 0; i < files.length; i++) {
                const reader = new FileReader();
                reader.onload = function (event) {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetNames = workbook.SheetNames;
                    
                    const workbookData = sheetNames.map(sheetName => {
                        return {
                            workbookName: sheetName,
                            sheets: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 })
                        };
                    });
                    workbooksData.push(workbookData);
                    
                    if (i === 0) {
                        // If it's the first file, initialize dropdowns and buttons
                        updateDropdowns();
                        updatePreview();
                    }
                };
                reader.readAsArrayBuffer(files[i]);
            }
        } else {
            // Handle the case when no files are selected
            previewArea.innerHTML = '<p>No files selected</p>';
        }
    });

    function updatePreview() {
        // Clear previous content
        while (previewArea.lastChild !== previewArea.firstChild) {
            previewArea.removeChild(previewArea.lastChild);
        }

        if (workbooksData.length > 0) {
            // Add the file and workbook dropdowns and buttons to the preview area
            previewArea.appendChild(fileDropdown);
            if (workbooksData[currentFileIndex].length > 1) {
                // Only add the workbook dropdown if there are multiple workbooks in the selected file
                previewArea.appendChild(workbookDropdown);
            }
            previewArea.appendChild(previousButton);
            previewArea.appendChild(nextButton);
        
            // Display the content of the selected workbook in a table
            const workbookDataToDisplay = workbooksData[currentFileIndex];
            const sheetData = workbookDataToDisplay[currentWorkbookIndex].sheets;
            const table = document.createElement('table');
            table.className = 'preview-table';
            for (let rowIndex = 0; rowIndex < Math.min(5, sheetData.length); rowIndex++) {
                const row = sheetData[rowIndex];
                const tr = document.createElement('tr');
            
                // Apply alternating row background colors
                if (rowIndex % 2 === 1) {
                    tr.className = 'alternate-row';
                }
            
                for (const cell of row) {
                    const td = document.createElement('td');
                    td.textContent = cell;
                    tr.appendChild(td);
                }
                table.appendChild(tr);
            }
            previewArea.appendChild(table);
        } else {
            // Handle the case when no workbooks are available
            previewArea.innerHTML += '<p>No workbooks available</p>';
        }
    }


    // Add logic to handle form submission
    document.getElementById('uploadForm').addEventListener('submit', function (event) {
        event.preventDefault(); // Prevent default form submission

        const formData = new FormData();
        for (const file of fileInput.files) {
            formData.append('file', file);
        }

        // Use fetch to send the form data to the server for processing
        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => response.text())
        .then(message => {
            alert(message); // Display the response from the server (you can customize this part)
        })
        .catch(error => {
            console.error('Error uploading file:', error);
        });
    });
});