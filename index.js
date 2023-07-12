const xlsx = require('xlsx');
const fs = require('fs-extra');

const sheets = [
  {
    name: 'Venituri',
    type: 'Venit',
    code_type: 'functional'
  },
  {
    name: 'Cheltuieli FCT',
    type: 'Cheltuiala',
    code_type: 'functional'
  },
  {
    name: 'Cheltuieli ECN',
    type: 'Cheltuiala',
    code_type: 'economic'
  }
]

function xlsxToJson (fileName, sheetName) {
  const workbook = xlsx.readFile(fileName);
  const sheet = workbook.Sheets[sheetName];
  const json = xlsx.utils.sheet_to_json(sheet, { header: 1, blankrows: true, raw: true, defval: null });
  return json;
}

function buildClasificatieDictionary (sheets) {
  const fileName = 'clasificatie.xlsx';
  const dictionary = new Map();
  sheets.forEach(sheet => {
    const json = xlsxToJson(fileName, sheet.name);

    // get header row index by finding the first row that contains the cell with the text 'Denumire'
    const headerRowIndex = json.findIndex(row => row.some(cell => cell === 'Denumire'));
    // get the column index from the header row by finding the cell that contains the text 'Denumire'
    const denumireIndex = json[headerRowIndex].findIndex(cell => cell === 'Denumire');
    // get the column index in the header row by finding the first cell that contains the text 'cod. Ec'
    const codIndex = json[headerRowIndex].findIndex(cell => cell === 'cod. Ec');

    // get data row index by finding the first row after the header row that contains something for the column found in codIndex
    const dataRowIndex = json.findIndex((row, index) => index > headerRowIndex && ![null, undefined, ''].includes(row[codIndex]));
    // get the data rows by slicing the json array from the data row index and filtering out empty rows
    const dataRows = json.slice(dataRowIndex).filter(row => row[codIndex] && row[denumireIndex]);

    // extract the data from the data rows by parsing each row and extracting the cod - denumire pairs
    const data = dataRows.map(row => {
      // extract the code
      let cod = String(row[codIndex]);
      // if sheet is Cheltuieli FCT then remove .xx from the code if starting with the 3rd character
      if (sheet.name === 'Cheltuieli FCT') {
        cod = cod.slice(0, 2) + cod.slice(5);
      }

      // eliminate all period characters '.' from cod
      cod = cod.replace(/\./g, '').padEnd(6, '0');
      const denumire = row[denumireIndex];
      return [cod, denumire];
    });

    // add to the dictionary

    if (dictionary.has(sheet.type)) {
      // if the dictionary already has the sheet type as a key, then add the data to the existing map
      const existingMap = dictionary.get(sheet.type);
      existingMap.set(sheet.code_type, new Map(data));
    } else {
      // create a new map with the sheet code type as the key and the data as the value
      dictionary.set(
        sheet.type,
        new Map([[sheet.code_type, new Map(data)]])
      );
    }
  });

  return dictionary;
}

function extractCodesFromDataFile (fileName) {
  const json = xlsxToJson(fileName, 'Sheet1');
  // find header row index by finding the first row that contains the cell with the text 'Tip Indicator'
  const headerRowIndex = json.findIndex(row => row.some(cell => cell === 'Tip Indicator'));
  // get the column index for 'Tip Indicator' from the header row by finding the cell that contains the text 'Tip Indicator'
  const tipIndicatorIndex = json[headerRowIndex].findIndex(cell => cell === 'Tip Indicator');
  // get the column index for 'Cod Functional' from the header row by finding the cell that contains the text 'Clasificatie Functionala'
  const codFunctionalIndex = json[headerRowIndex].findIndex(cell => cell === 'Clasificatie Functionala');
  // get the column index for 'Cod Functional Descriere' from the header row by finding the cell that contains the text 'Clasificatie Functionala Descriere'
  const codFunctionalDescriereIndex = json[headerRowIndex].findIndex(cell => cell === 'Clasificatie Functionala Descriere');
  // get the column index for 'Cod Economic' from the header row by finding the cell that contains the text 'Clasificatie Economica'
  const codEconomicIndex = json[headerRowIndex].findIndex(cell => cell === 'Clasificatie Economica');
  // get the column index for 'Cod Economic Descriere' from the header row by finding the cell that contains the text 'Clasificatie Economica Descriere'
  const codEconomicDescriereIndex = json[headerRowIndex].findIndex(cell => cell === 'Clasificatie Economica Descriere');

  // get data row index by finding the first row after the header row that contains something for the column found in tipIndicatorIndex
  const dataRowIndex = json.findIndex((row, index) => index > headerRowIndex && ![null, undefined, ''].includes(row[tipIndicatorIndex]));
  // get the data rows by slicing the json array from the data row index and filtering out empty rows and total rows (totals don't have codes)
  const dataRows = json.slice(dataRowIndex).filter(row => row[codFunctionalIndex]);

  // build a dictionary of codes
  const codes = new Map();
  dataRows.forEach(row => {
    const tipIndicator = row[tipIndicatorIndex].trim();
    const codFunctional = row[codFunctionalIndex];
    const codFunctionalDescriere = row[codFunctionalDescriereIndex];
    const codEconomic = row[codEconomicIndex];
    const codEconomicDescriere = row[codEconomicDescriereIndex];

    if (codes.has(tipIndicator)) {
      const existingMap = codes.get(tipIndicator);
      // get the functional map
      const functionalMap = existingMap.get('functional');
      // add the functional code to the functional map
      functionalMap.set(codFunctional, codFunctionalDescriere);

      if (tipIndicator !== 'Venit') {
        // get the economic map
        const economicMap = existingMap.get('economic');
        // add the economic code to the economic map
        economicMap.set(codEconomic, codEconomicDescriere);
      }
    } else {
      // create a new map with the functional and economic codes
      if (tipIndicator === 'Venit') {
        codes.set(
          tipIndicator,
          new Map([
            ['functional', new Map([[codFunctional, codFunctionalDescriere]])]
          ])
        );
      } else {
        codes.set(
          tipIndicator,
          new Map([
            ['functional', new Map([[codFunctional, codFunctionalDescriere]])],
            ['economic', new Map([[codEconomic, codEconomicDescriere]])]
          ])
        );
      }
    }
  });

  return codes;
}

function main () {
  // build clasificatie dictionary
  const clasificatieDictionary = buildClasificatieDictionary(sheets);

  // get a list of data files in the ./files folder and filter out the lock files '.~lock.*'
  const files = fs.readdirSync('./files').filter(file => !file.startsWith('.~lock.'));
  const missingCodes = new Map();

  files.forEach((file, index) => {
    console.log(`Processing ${file} (${index + 1} of ${files.length})`);
    // extract the identifiers from the file name
    // 20230630_FXB-EXB-901_TREZ002_4562150_4562150_02_51422020.xlsx
    // yyyymmdd_FXB-EXB-901_TREZ002_identifier1_identifier2_xy_xxxxxxxx.xlsx
    const identifiers = file.split('_').slice(3, 5).join('_');

    // extract codes from the first data file
    const codes = extractCodesFromDataFile(`./files/${file}`);

    // compare the dictionary with the extracted codes from the data file to see if there are any extra codes in the data file
    codes.forEach((data, tipIndicator) => {
      data.forEach((coduri, tipCod) => {
        const dictionary = clasificatieDictionary.get(tipIndicator);
        if (dictionary) {
          const clasificatie = dictionary.get(tipCod);
          if (clasificatie) {
            coduri.forEach((descriere, cod) => {
              if (!clasificatie.has(cod)) {
                if (missingCodes.has(identifiers)) {
                  const existingMap = missingCodes.get(identifiers);
                  // if there's a set for the tipIndicator then add the cod to the set
                  if (existingMap.has(tipIndicator)) {
                    const existingSet = existingMap.get(tipIndicator);
                    existingSet.add({
                      tipCod,
                      cod,
                      descriere
                    });
                  } else {
                    existingMap.set(tipIndicator, new Set([{
                      tipCod,
                      cod,
                      descriere
                    }]));
                  }
                } else {
                  missingCodes.set(identifiers, new Map([[tipIndicator, new Set([{
                    tipCod,
                    cod,
                    descriere
                  }])]]));
                }
              }
            });
          } else {
            console.log(`Missing ${tipCod} map`);
          }
        } else {
          console.log(`Missing ${tipIndicator} map`);
        }
      });
    });
  });

  // transform the missingCodes Map and the containing sets into an array of objects
  const missingCodesArrayOfObjects = [];
  missingCodes.forEach((map, identifiers) => {
    map.forEach((set, tipIndicator) => {
      set.forEach(cod => {
        const [cifOp, cifOs] = identifiers.split('_');
        missingCodesArrayOfObjects.push({
          cifOp,
          cifOs,
          tipIndicator,
          ...cod
        });
      });
    });
  });

  // write the missing codes to an xlsx file
  const workbook = xlsx.utils.book_new();
  const sheet = xlsx.utils.json_to_sheet(missingCodesArrayOfObjects);
  xlsx.utils.book_append_sheet(workbook, sheet, 'Missing Codes');
  xlsx.writeFile(workbook, 'missing_codes.xlsx');
}

main();