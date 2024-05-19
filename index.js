const oracledb = require('oracledb');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function exportTablesToExcel(user, password, connectString, outputPath) {
    let connection;

    try {
        connection = await oracledb.getConnection({
            user: user,
            password: password,
            connectString: connectString
        });

        const result = await connection.execute(`
            SELECT table_name
            FROM all_tables
            WHERE owner = upper('${user}')
            ORDER BY table_name DESC
        `);

        const tables = result.rows.map(row => row[0]);
        const today = new Date().toISOString().split('T')[0];
        const directory = path.join(outputPath, today);

        if (!fs.existsSync(directory)) {
            fs.mkdirSync(directory);
        }

        for (const tableName of tables) {
            const data = await connection.execute(`SELECT * FROM ${tableName}`);
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet(tableName);
            const columns = data.metaData.map((meta) => ({ header: meta.name, key: meta.name }));
            worksheet.columns = columns;
            worksheet.getRow(1).font = { bold: true };
            worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getRow(1).eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFCCCCCC' }
                };
            });

            data.rows.forEach(row => {
                const rowData = {};
                row.forEach((value, index) => {
                    rowData[columns[index].key] = value;
                });
                const addedRow = worksheet.addRow(rowData);
                addedRow.eachCell((cell) => {
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                });
            });

            worksheet.columns.forEach(column => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, cell => {
                    const cellValue = cell.value ? cell.value.toString() : '';
                    if (cellValue.length > maxLength) {
                        maxLength = cellValue.length;
                    }
                });
                column.width = maxLength + 2;
            });

            const filePath = path.join(directory, `${tableName}.xlsx`);
            await workbook.xlsx.writeFile(filePath);

            console.log(`Data from table ${tableName} has been written to ${filePath}`);
        }
    } catch (err) {
        console.error(err);
    } finally {
        if (connection) {
            try {
                await connection.close();
            } catch (err) {
                console.error(err);
            }
        }
    }
}

const user = process.argv[2];
const password = process.argv[3];
const connectString = process.argv[4];
const outputPath = process.argv[5];

exportTablesToExcel(user, password, connectString, outputPath);
 