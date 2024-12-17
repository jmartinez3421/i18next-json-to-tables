const fs = require('fs/promises');
const path = require('path');
const ExcelJs = require('exceljs');
const inquirer = require('inquirer');
require("colors");

const resultsDir = './results';

const getExcelFiles = async (excelDir, extension = null) => {
    try {
        const files = await fs.readdir(excelDir);

        const excelFiles = [];
        for (const file of files) {
            const filePath = path.join(excelDir, file); // Obtener la ruta completa
            const stats = await fs.stat(filePath); // Obtener información del fichero

            if (stats.isFile()) { // Comprobar si es un fichero
                if (extension) { // Si se especificó una extensión
                    if (path.extname(file).toLowerCase() === extension.toLowerCase()) {
                        excelFiles.push(file);
                    }
                } else {
                    excelFiles.push(file); // Si no hay extensión, añadir todos los ficheros
                }
            }
        }
        return excelFiles;
    } catch (error) {
        console.error(`Error al leer el directorio ${excelDir}:`, error);
        return []; // Devolver un array vacío en caso de error
    }
};

async function readFileAsArrayBuffer(filePath) {
    try {
        const fileBuffer = await fs.readFile(filePath);
        return fileBuffer.buffer; // Obtener el ArrayBuffer del Buffer
    } catch (error) {
        console.error(`Error al leer el archivo ${filePath}:`, error);
        return null;
    }
}

/**
 * Ask the user for a string
 * @param message {string}
 * @returns {Promise<number>}
 */
const askNumber = async (message) => {
    const {number} = await inquirer.prompt([
        {
            type: "input",
            name: "number",
            message,
        },
    ]);
    if (isNaN(number)) return askNumber(message);
    return Number(number);
}


const processExcelFile = async (filePath, fileName) => {
    try {
        const numberOfLanguages = await askNumber("How many languages do you have?");
        readFileAsArrayBuffer(filePath).then(async (buffer) => {
            try {
                const workbook = new ExcelJs.Workbook();
                const wbFile = await workbook.xlsx.load(buffer);
                const worksheets = wbFile.worksheets;

                if (worksheets.length === 0) throw new Error("No sheet found");

                const worksheetTranslations = {};
                for (const worksheet of worksheets) {
                    console.log(`Processing sheet ${worksheet.name}...`);
                    const languages = [];
                    const translations = {};
                    worksheet.eachRow((row, rowNumber) => {
                        const key = row.getCell(1).value?.toString() || "";

                        if (rowNumber === 1) {
                            for (let i = 2; i <= numberOfLanguages + 1; i++) {
                                const cell = row.getCell(i);
                                if (cell.value) {
                                    languages.push(cell.value.toString());
                                    translations[cell.value.toString()] = {};
                                }
                            }
                            return;
                        }

                        if (key) {
                            for (let i = 2; i <= numberOfLanguages + 1; i++) {
                                const cell = row.getCell(i);
                                if (key.includes(".")) {
                                    const keyParts = key.split(".");
                                    if (keyParts.length === 2) {
                                        if (!translations[languages[i - 2]][keyParts[0]]) {
                                            translations[languages[i - 2]][keyParts[0]] = {};
                                        }
                                        translations[languages[i - 2]][keyParts[0]][keyParts[1]] = cell.value?.toString() || "";
                                    } else if (keyParts.length === 3) {
                                        if (!translations[languages[i - 2]][keyParts[0]]) {
                                            translations[languages[i - 2]][keyParts[0]] = {};
                                        }
                                        if (!translations[languages[i - 2]][keyParts[0]][keyParts[1]]) {
                                            translations[languages[i - 2]][keyParts[0]][keyParts[1]] = {};
                                        }
                                        translations[languages[i - 2]][keyParts[0]][keyParts[1]][keyParts[2]] = cell.value?.toString() || "";
                                    } else {
                                        translations[languages[i - 2]][key] = cell.value?.toString() || "";
                                    }
                                } else {
                                    translations[languages[i - 2]][key] = cell.value?.toString() || "";
                                }
                            }
                        }
                    });
                    worksheetTranslations[worksheet.name] = translations;

                    const resultsFiles = await fs.readdir(resultsDir);
                    if (!resultsFiles.includes(fileName)) {
                        await fs.mkdir(path.join(resultsDir, fileName));
                    } else {
                        await fs.rmdir(path.join(resultsDir, fileName), {recursive: true});
                        await fs.mkdir(path.join(resultsDir, fileName));
                    }

                    const resultDirPath = path.join(resultsDir, fileName);

                    for (const language of languages) {
                        await fs.mkdir(path.join(resultDirPath, language));
                    }

                    const filesToWrite = Object.keys(worksheetTranslations);
                    for (const file of filesToWrite) {
                        const fileTranslations = worksheetTranslations[file];
                        Object.entries(fileTranslations).forEach(([language, translations]) => {
                            const filePath = path.join(resultDirPath, language, `${file}.json`);

                            const translationsKeys = Object.keys(translations).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                            const translationsSorted = {};
                            for (const key of translationsKeys) {
                                translationsSorted[key] = translations[key];
                            }

                            const fileTranslationsAsString = JSON.stringify(translationsSorted, null, 2);
                            fs.writeFile(filePath, fileTranslationsAsString);
                        });
                    }
                }
            } catch (error) {
                console.error(`Error al procesar el fichero ${filePath}:`, error);
            }
        });
    } catch (error) {
        console.error(`Error al leer el fichero ${filePath}:`, error);
    }
};

const startConversion = async () => {
    const excelDir = './excels';
    try {
        const xlsxFiles = await getExcelFiles(excelDir, '.xlsx');
        for (const file of xlsxFiles) {
            await processExcelFile(path.join(excelDir, file), file);
        }
    } catch (error) {
        console.error("Error en startConversion:", error)
    }
};

startConversion();
