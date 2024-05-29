const {readdir} = require('fs/promises');
const {readFileSync, writeFileSync} = require('fs');
const path = require('path');
const ExcelJs = require('exceljs');

// Indicate the folder where the translations are stored
const translationsDir = './translations';

// Indicate the folder where the results will be stored
const resultsDir = './results';

// Include all the languages that will be used.
// The keys are the names of the languages folders in the translations folder,
// and the values are the names that will be displayed in the result files.
const languagesMap = {
    'es': 'Español',
    'en': 'Inglés',
    'pt-BR': 'Brasileiro',
};

const languages = Object.keys(languagesMap);
const headers = ["Key", ...Object.values(languagesMap)];

// Each translation file will have a key, and the value will have all the keys and their translations in the different languages
// Example: {"common": { "hello": ["Hola", "Olá", "你好"] }}
const translations = {};

/**
 * Set the value of a translation. If you add new languages, you need to add a new case here.
 * The position of the language in the array has to match the position of the language in the languagesMap.
 * @param language {string}
 * @param file {string}
 * @param key {string}
 * @param value {string}
 */
const setValue = (language, file, key, value) => {
    if (!translations[file][key]) {
        translations[file][key] = [];
    }
    switch (language) {
        case 'es':
            translations[file][key][0] = value.trim();
            break;
        case 'en':
            translations[file][key][1] = value.trim();
            break;
        case 'pt-BR':
            translations[file][key][2] = value.trim();
            break;
        default:
            break;
    }
}

/**
 * Reads an object and sets the value of the translations.
 * <br />
 * If the object is an object, it will recursively call itself to read the nested objects.
 * The key of the object will include the previous key and the key of the object itself.
 * This is done to create a nested structure in the translations object.
 * @param language
 * @param file
 * @param prevKey
 * @param object
 */
const readObject = (language, file, prevKey, object) => {
    if (object instanceof Object) {
        Object.entries(object).forEach(([key, value]) => {
            if (value instanceof Object) {
                readObject(language, file, `${prevKey}.${key}`, value);
            } else {
                setValue(language, file, `${prevKey}.${key}`, value);
            }
        });
    } else {
        return setValue(language, file, prevKey, object);
    }
}

/**
 * Generates an Excel file with all the translations. Each translation file will be a worksheet in the Excel file.
 * @param data {{ [key: string]: [string[]] }}
 */
const generateExcel = (data) => {
    const workbook = new ExcelJs.Workbook();

    Object.entries(data).forEach(([sheetName, values]) => {
        const worksheet = workbook.addWorksheet(sheetName, {
            pageSetup: {
                horizontalCentered: true,
                verticalCentered: true,
            },
        });

        // Add a header row with the column names
        worksheet.addRow(headers);

        // Add the translations to the worksheet
        values.forEach((row) => {
            worksheet.addRow(row);
        });

        const numberOfColumns = headers.length;

        // Set row height and alignment
        for (let i = 0; i <= values.length + 1; i++) {
            const row = worksheet.getRow(i);
            row.height = 20;
            row.alignment = {
                vertical: "middle",
            };
            row.font = {
                name: "Arial",
                size: 12,
            };
            row.eachCell((cell) => {
                cell.border = {
                    bottom: { style: "thin", color: { argb: "303030" } },
                    top: { style: "thin", color: { argb: "303030" } },
                    right: { style: "thin", color: { argb: "303030" } },
                    left: { style: "thin", color: { argb: "303030" } },
                };
            });
        }

        // Set column 1 bold
        worksheet.getColumn("A").font = {
            bold: true,
            name: "Arial",
            size: 12,
        };

        const columnsMaxLength = {};
        for (let i = 0; i < numberOfColumns; i++) {
            // Get max length of the column
            columnsMaxLength[String.fromCharCode(65 + i)] = Math.max(...values.map((row) => {
                if (row[i]) return row[i].length;
                return 0;
            }));

            // Set column header style
            const cell = worksheet.getCell(`${String.fromCharCode(65 + i)}1`);
            cell.font = {
                bold: true,
                color: { argb: "4c7cb2" },
                name: "Arial",
                size: 12,
            };
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "d9e1f2" },
            };
            cell.alignment = {
                horizontal: "center",
                vertical: "middle",
            };
        }

        // Set column width to the maximum length of the column
        Object.keys(columnsMaxLength).forEach((key) => {
            worksheet.getColumn(key).width = columnsMaxLength[key] + 10;
        });
    })

    // Generate the Excel file
    workbook.xlsx.writeBuffer().then((excelData) => {
        writeFileSync(`${resultsDir}/translations.xlsx`, excelData);
        console.log(`Excel with all translations has been generated in ${resultsDir}/translations.xlsx`);
    });

}

/**
 * Reads the translations from the translations folder and converts them to CSV files and Excel file
 * @returns {Promise<void>}
 */
const readTranslations = async () => {
    // Loop through all the languages folders
    for (const language of languages) {
        console.log(`Reading ${language} translations...`);
        try {
            // Get the names of all the files in the language folder
            const files = await readdir(`${translationsDir}/${language}`);

            // Get the values of all the files in the language folder and map them into the translations object
            files.forEach(file => {
                console.log(`\tReading ${file} file...`);

                // If the file has no key on the translations object, create it
                if (!translations[file]) {
                    translations[file] = {};
                }

                // Get the path of the file
                const filePath = path.join(translationsDir, language, file);
                try {
                    // Read the file as a JSON object
                    const fileTranslations = JSON.parse(readFileSync(filePath, 'utf8'));

                    // Map all the values of the json object to the translations object
                    Object.entries(fileTranslations).forEach(([key, value]) => {
                        readObject(language, file, key, value);
                    });
                } catch (err) {
                    console.error(`Error while reading ${filePath}:`, err);
                }
            });
        } catch (error) {
            console.log(error);
        }
    }

    // Map with the name of the file as the key and an array of arrays with the keys and their translations in the different languages
    const excelObjects = {};

    Object.entries(translations).forEach(([file, translations]) => {
        console.log(`Writing ${file} file...`);

        // Get the name of the file without the extension
        const resultFileName = file.replace(".json", "");

        // Initialize the csv data with the headers
        const csvData = [headers];
        for (const key in translations) {
            // Add the key and its translations to the csv data
            csvData.push([
                key,
                    ...translations[key].map((translation) =>`"${translation}"`)
            ]);

            // If the csvObject doesn't have the resultFileName key, create it
            if (!excelObjects[resultFileName]) {
                excelObjects[resultFileName] = [];
            }

            // Add the key and its translations to the csvObject
            excelObjects[resultFileName].push([
                key,
                ...translations[key]
            ]);
        }

        // Convert the csv data to a string
        const csvAsString = csvData.join('\n').replaceAll("undefined", "");

        // Write the csv data to a file
        writeFileSync(`${resultsDir}/${resultFileName}.csv`, csvAsString);
    });

    // Generate the Excel file with all the translations
    generateExcel(excelObjects);

    console.log(`All translations have been converted to CSV. You can find them in the ${resultsDir} folder.`);
}

readTranslations();
