import pkg from 'xlsx';
const { readFile, utils, writeFile } = pkg
// Helper function to split text by numbered statements (e.g., 1., 2., 10.)

const userEmail = 'Avinash Chauhan <avinashchauhan@gofynd.com>';
const path = `CatalogCloud`;

// Usage example
const inputFilePath = 'excelFiles/ADO Test.xlsx';   // Input Excel file path
const outputFilePath = 'excelFiles/ADO Test_mod.xlsx'; // Output Excel file path
const columnsToModify = [4, 6];         // Indices of columns to modify (e.g., C = 2, D = 3)

// Function to read from one Excel file, modify data from two columns, and write to another Excel file
function readAndModifyExcel(inputFilePath, outputFilePath, columnsToModify) {
    const workbook = readFile(inputFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = utils.sheet_to_json(worksheet, { header: 1 });
    const newHeader = ['ID', 'Work Item Type', 'Title', 'Test Step', 'Step Action', 'Step Expected', 'Area Path', 'Assigned To', 'State'];

    let modifiedData = [];
    let row = [];

    for (let i = 1; i < data.length; i++) {
        let columnSplits = [];
        row = data[i];
        if (row[0] == null) {
            break;
        }
        columnsToModify.forEach((columnIndex) => {
            const cellValue = row[columnIndex];
            if (cellValue && typeof cellValue === 'string') {
                const splitStatements = splitNumberedStatements(cellValue);
                columnSplits.push(splitStatements);
            } else {
                columnSplits.push([cellValue]); // If no split, keep original value
            }
        });
        const newRows = mergeColumnData(row, columnSplits, columnsToModify);
        modifiedData.push(...newRows);
    }

    // If there is no data after modifications, add an empty row (for debugging purposes)
    if (modifiedData.length === 0) {
        console.log("Warning: No data to write.");
    }
    modifiedData.unshift(newHeader)
    const newWorksheet = utils.aoa_to_sheet(modifiedData);
    const newWorkbook = utils.book_new();
    utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
    writeFile(newWorkbook, outputFilePath);

    console.log('Excel file has been modified and saved to', outputFilePath);
}

function splitNumberedStatements(text) {
    const result = [];
    let currentStatement = '';
    let regex = /^\d+\.\s/;  // Looks for numbers followed by a period and a space

    // Split the text by newlines or similar breaks
    let lines = text.split(/\r?\n/);

    for (let line of lines) {
        // If the line starts with a numbered statement (e.g., "1. ")
        if (regex.test(line.trim())) {
            if (currentStatement.trim()) {
                result.push(currentStatement.trim());
            }
            currentStatement = line.trim(); // Start a new statement
        } else {
            currentStatement += ' ' + line.trim(); // Append to the current statement
        }
    }

    // Push the last statement if any
    if (currentStatement.trim()) {
        result.push(currentStatement.trim());
    }

    return result;
}

// Function to merge split column data into the same row structure
function mergeColumnData(originalRow, columnSplits, columnIndices) {
    const maxRows = Math.max(...columnSplits.map(split => split.length));
    const newRows = [];

    for (let i = 0; i <= maxRows; i++) {
        const newRow = new Array(originalRow.length).fill(null); // Fill with null initially

        // Copy the original data for columns that aren't being modified
        originalRow.forEach((cell, colIdx) => {

            if (colIdx != 0 && (!columnIndices.includes(colIdx))) {
                if (colIdx == 2) {
                    newRow[colIdx] = i === 0 ? cell : null;
                    newRow[1] = i === 0 ? `Test Case` : null;
                    newRow[6] = i === 0 ? path : null;
                    newRow[7] = i === 0 ? userEmail : null;
                    newRow[8] = i === 0 ? 'Design' : null;
                }
            }
        });

        // Fill in the split column data
        columnSplits.forEach((split, idx) => {
            let colIndex = idx === 0 ? columnIndices[idx] : columnIndices[idx] - 1;
            if (i <= split.length) {
                newRow[3] = i !== 0 ? i : null;
                newRow[colIndex] = i !== 0 ? split[i - 1] : null;

            }
        });
        newRows.push(newRow);
    }
    return newRows;
}

readAndModifyExcel(inputFilePath, outputFilePath, columnsToModify);

console.log('Excel file has been modified and saved to', outputFilePath);