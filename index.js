const Excel = require('exceljs');
const createCsvWriter = require('csv-writer').createArrayCsvWriter;

async function convertExcelToCSV() {
    // Get the Excel file name from command line arguments
    const excelFileName = process.argv[2];

    // Create a new workbook
    const workbook = new Excel.Workbook();

    // Read the Excel file
    await workbook.xlsx.readFile(excelFileName);

    // Get the first worksheet
    const worksheet = workbook.getWorksheet(1);

    // Prepare data for CSV
    const data = [];
    let rowCount = 0;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        rowCount++;
        if (rowCount <= 6) return;
        const rowData = row.values.slice(1);  // slice(1) to remove the leading empty value
        data.push(rowData);
    });

    // Define CSV writer
    const csvWriter = createCsvWriter({
        path: 'output.csv',
    });

    // Write data to CSV
    await csvWriter.writeRecords(data);
}

convertExcelToCSV();
