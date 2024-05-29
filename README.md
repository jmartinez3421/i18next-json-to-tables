# i18next JSON to Table

This is a simple tool that converts the translation files from the i18next JSON format to a table.

It will generate a CSV file for each translation file with a column for the key and a column for each value for that key in the different languages.

It will also generate an Excel file with the same data,
but in a more readable format and instead of creating a file for each translation file,
it will create a single file with a sheet for each translation file.

## Usage
First you need to install the dependencies:

```bash
pnpm install
```

Then you have to place the translation files in the `i18n` folder. The translation files should be inside a folder for each language.

For example:

- translations
  - en
    - common.json
    - errors.json
  - de
    - common.json
    - errors.json

After that you can run the script:

```bash
pnpm start
```

You will be asked for a label for each language, after that the script will start converting the files.

> If you want to use the folder name as label leave it empty and press enter.

When the script is finished, you will find all the results in the `results` folder.
