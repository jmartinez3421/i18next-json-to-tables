const fs = require('fs/promises');
const path = require('path');
const ExcelJs = require('exceljs');
const inquirer = require('inquirer');
require("colors");

const resultsDir = './results'; // Directory to store the results
const excelDir = './excels'; // Directory where the excel files are stored

/**
 * Get all excel files in the directory defined by excelDir
 * @param extension
 * @return {Promise<*[]>}
 */
const getExcelFiles = async (extension = null) => {
    try {
        const files = await fs.readdir(excelDir); // Get all the files in the directory

        const excelFiles = [];
        for (const file of files) {
            const filePath = path.join(excelDir, file);
            const stats = await fs.stat(filePath);

            // If it's a file and has the specified extension add it to the list
            if (stats.isFile()) {
                if (extension) {
                    if (path.extname(file).toLowerCase() === extension.toLowerCase()) {
                        excelFiles.push(file);
                    }
                } else {
                    excelFiles.push(file);
                }
            }
        }
        return excelFiles;
    } catch (error) {
        console.error(`Error reading directory ${excelDir}:`, error);
        return [];
    }
};

/**
 * Read a file as an ArrayBuffer
 * @param filePath
 * @return {Promise<ArrayBuffer|null>}
 */
async function readFileAsArrayBuffer(filePath) {
    try {
        const fileBuffer = await fs.readFile(filePath);
        return fileBuffer.buffer;
    } catch (error) {
        console.error(`Error reading file ${filePath}:`, error);
        return null;
    }
}

/**
 * Ask the user for a number
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

/**
 * Process all the worksheets in the given excel file and write the translations on i18n-json files in the results directory
 * @param filePath
 * @param fileName
 * @return {Promise<void>}
 */
const processExcelFile = async (filePath, fileName) => {
    try {
        // Ask the user for the number of languages
        const numberOfLanguages = await askNumber("How many languages do you have?");

        readFileAsArrayBuffer(filePath).then(async (buffer) => {
            try {
                // Load the excel file as a workbook
                const workbook = new ExcelJs.Workbook();
                const wbFile = await workbook.xlsx.load(buffer);

                // Get all the worksheets in the workbook
                const worksheets = wbFile.worksheets;

                if (worksheets.length === 0) throw new Error("No sheet found");

                const worksheetTranslations = {};
                for (const worksheet of worksheets) {
                    console.log(`Processing sheet ${worksheet.name}...`);
                    const languages = []; // Array to store the languages
                    const translations = {}; // Object to store the translations

                    // Iterate over each row in the worksheet
                    worksheet.eachRow((row, rowNumber) => {
                        const key = row.getCell(1).value?.toString() || ""; // Get the key from the first column

                        // If the row number is 1, get the number of languages and initialize the translations object
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

                        // If the key is not empty, iterate over each language and add the translations to the translations object
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
                    // Add the translations object to the worksheet attribute of the worksheetTranslations object
                    worksheetTranslations[worksheet.name] = translations;

                    // Create the results directories if they don't exist
                    const resultsFiles = await fs.readdir(resultsDir);
                    if (!resultsFiles.includes(fileName)) {
                        await fs.mkdir(path.join(resultsDir, fileName));
                    } else {
                        await fs.rmdir(path.join(resultsDir, fileName), {recursive: true});
                        await fs.mkdir(path.join(resultsDir, fileName));
                    }

                    // Get the path to the results directory
                    const resultDirPath = path.join(resultsDir, fileName);

                    // Create the languages directories in the results directory
                    for (const language of languages) {
                        await fs.mkdir(path.join(resultDirPath, language));
                    }

                    // Get the list of files to write
                    const filesToWrite = Object.keys(worksheetTranslations);

                    // Iterate over the files to write
                    for (const file of filesToWrite) {
                        const fileTranslations = worksheetTranslations[file]; // Get the translations for the current file

                        // Iterate over the languages
                        Object.entries(fileTranslations).forEach(([language, translations]) => {
                            // Get the path to the file
                            const filePath = path.join(resultDirPath, language, `${file}.json`);

                            // Sort the keys alphabetically
                            const translationsKeys = Object.keys(translations).sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
                            const translationsSorted = {};
                            for (const key of translationsKeys) {
                                translationsSorted[key] = translations[key];
                            }

                            // Convert the object to a string and write it to the file
                            const fileTranslationsAsString = JSON.stringify(translationsSorted, null, 4);
                            fs.writeFile(filePath, fileTranslationsAsString);
                        });
                    }
                }
            } catch (error) {
                console.error(`Error processing file ${filePath}:`, error);
            }
        });
    } catch (error) {
        console.error(`Error reading file ${filePath}:`, error);
    }
};

/**
 * Start the conversion process from the excel translations to i18n-json files
 * @return {Promise<void>}
 */
const startConversion = async () => {
    try {
        const xlsxFiles = await getExcelFiles('.xlsx');
        for (const file of xlsxFiles) {
            await processExcelFile(path.join(excelDir, file), file);
        }
    } catch (error) {
        console.error("Error starting the conversion process:", error)
    }
};

startConversion();
