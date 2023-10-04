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
    workbookDropdown.style.marginLeft = '10px';
    workbookDropdown.addEventListener('change', function () {
        currentWorkbookIndex = workbookDropdown.selectedIndex;
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

        // If workbooksData is empty, return (no need to update workbook dropdown)
        if (workbooksData.length === 0) return;

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

    fileInput.addEventListener('change', function () {
        previewArea.innerHTML = ''; // Clear previous content
        workbooksData = []; // Clear previous workbook data

        const files = fileInput.files;
        if (files.length > 0) {
            let loadedFiles = 0; // Counter for tracking loaded files

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

                    loadedFiles++; // Increment the loadedFiles counter

                    if (loadedFiles === files.length) {
                        // If all files are loaded, initialize dropdowns and buttons
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

            // Check if the selected file has multiple workbooks
            const workbooks = workbooksData[currentFileIndex];
            if (workbooks.length > 1) {
                // Only show the workbook dropdown if there are multiple workbooks in the selected file
                previewArea.appendChild(workbookDropdown);
            } else {
                // Hide the workbook dropdown if there is only one workbook in the selected file
                workbookDropdown.innerHTML = '';
            }

            // Update the workbookDataToDisplay variable
            const workbookDataToDisplay = workbooksData[currentFileIndex];

            // Display the content of the selected workbook in a table
            const sheetData = workbookDataToDisplay[currentWorkbookIndex].sheets;
            const table = document.createElement('table');
            table.style.marginTop = '10px';
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

    // logic to dynamically fill up input field
    document.getElementById('fileInput').addEventListener('change', function() {
        const fileInput = document.getElementById('fileInput');
        const tableIdInput = document.getElementById('tableIdInput');
    
        if (fileInput.files.length > 0) {
        const uploadedFile = fileInput.files[0];
        const fileName = uploadedFile.name;
    
        // Extract the first word from the file name
        const firstWord = fileName.split(/[_\s]/)[0];
    
        // Get the current date in the format DD-MM-YY
        const currentDate = new Date().toLocaleDateString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: '2-digit'
        }).replace(/\//g, '-');
    
        // Construct the dynamic table ID
        const dynamicTableId = `${firstWord}_${currentDate}`;
    
        // Fill the table ID input field
        tableIdInput.value = dynamicTableId;
        }
    });


    // Add logic to handle form submission
    document.getElementById('uploadForm').addEventListener('submit', function (event) {
        event.preventDefault(); // Prevent default form submission

        const formData = new FormData();
        for (const file of fileInput.files) {
            formData.append('file', file);
        }

        // Add the value of the tableIdInput to the formData
        const tableIdInput = document.getElementById('tableIdInput').value;

        formData.append('tableIdInput', tableIdInput);

        fetch('/upload', {
            method: 'POST',
            body: {
                file: fileInput,
                tableIdInput: tableIdInput
            }
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
