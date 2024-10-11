
const dragArea = document.getElementById('dragArea');
const fileInput = document.getElementById('excelFile');
const fileList = document.getElementById('fileList');

['dragenter', 'dragover'].forEach(eventType => {
    dragArea.addEventListener(eventType, (e) => {
        e.preventDefault();
        dragArea.classList.add('dragging');
    });
});

['dragleave', 'drop'].forEach(eventType => {
    dragArea.addEventListener(eventType, (e) => {
        e.preventDefault();
        dragArea.classList.remove('dragging');
    });
});

dragArea.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFiles(files);
    }
});

dragArea.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
        handleFiles(fileInput.files);
    }
});

// Function to handle multiple files
function handleFiles(files) {
    for (let file of files) {
        processFile(file);
    }
}

function processFile(file) {
    const listItem = document.createElement('li');
    listItem.innerHTML = `
      <span>${file.name}</span>
      <div class="file-status">
        <div class="loading-spinner"></div>
        <span class="status-text processing">Elaborazione...</span>
      </div>
    `;
    fileList.appendChild(listItem);

    // Read the file as binary
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Get the first worksheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Extract the specified columns (0 and 2 in this case)
        const columns = [0, 2]; // Indices of the columns you want to extract
        const contentColumn = []; // Array to store the extracted column data
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Initialize contentColumn to have as many arrays as there are columns to be extracted
        columns.forEach(() => contentColumn.push([])); // Prepare nested arrays for each column

        // Process each specified column
        columns.forEach(col => {
            for (let row = range.s.r + 1; row <= range.e.r; row++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = worksheet[cellAddress];
                if (cell) {
                    let cellExtracted = cell.v;
                    if (col === 0) {
                        cellExtracted = cellExtracted.split(' ').slice(0, -1).join(' ');
                    }
                    else if (col === 2) {
                        cellExtracted = cellExtracted.slice(1).charAt(0).toUpperCase() + cellExtracted.slice(1);
                    }
                    contentColumn[columns.indexOf(col)].push(cellExtracted);
                }
            }
        });

        // Ensure all columns have the same length by filling with empty strings
        const maxRows = Math.max(...contentColumn.map(col => col.length));
        const finalContent = [];

        for (let i = 0; i < maxRows; i++) {
            const row = columns.map((col, colIndex) => contentColumn[colIndex][i] || ""); // Fill with empty strings if undefined
            finalContent.push(row);
        }

        // Create a new worksheet with the extracted columns
        const newWorksheet = XLSX.utils.aoa_to_sheet(finalContent);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'ProcessedData');

        // Generate a new Excel file
        const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

        // Create a download link for the processed file
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const downloadLink = document.createElement('a');
        const url = window.URL.createObjectURL(blob);
        downloadLink.href = url;
        downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + '_processed.xlsx';
        downloadLink.classList.add('btn', 'btn-success', 'btn-sm');
        downloadLink.textContent = 'Scarica';

        // Update file status to "Complete" and append download link
        listItem.querySelector('.file-status').innerHTML = `<span class="status-text complete">Elaborato</span>`;
        listItem.appendChild(downloadLink);

        // Revoke the object URL after download to free memory
        downloadLink.addEventListener('click', () => {
            setTimeout(() => window.URL.revokeObjectURL(url), 100);
        });
    };

    // Read the file as binary string
    reader.readAsArrayBuffer(file);
}



