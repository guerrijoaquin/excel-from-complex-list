import ExcelJS from 'exceljs';

enum Movements {
  RIGTH = 'RIGTH',
  LEFT = 'LEFT',
  TOP = 'TOP',
  BOTTOM = 'BOTTOM',
}

enum Actions {
  MOVE = 'MOVE',
  WRITE = 'WRITE',
}

export interface OneSheet {
  name: string;
  data: any;
}

const sheetConfig = { views: [{ showGridLines: true }] };
const borderTopStyle: Partial<ExcelJS.Borders> = {
  top: {
    color: {
      argb: '000000',
    },
    style: 'thin',
  },
};
let lastAction: Actions = Actions.WRITE;
let lastRow: number = 0;
let headerFinalRow: number = 0;

/**
 * Async excel file generation
 * @param filename Name of the file to be written
 * @param sheets sheets of the book
 * @returns the file
 */
export const getExcelFromJSONList = async (filename: string, sheets: OneSheet[]) => {
  return true;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Tarjeta Prepaga | Comafi';

  sheets.forEach(({ name, data }) => renderSheet(workbook, name, data));

  return await workbook.xlsx.writeBuffer({ filename: `${filename}.xlsx` });
};

const renderSheet = (wb: ExcelJS.Workbook, name: string, data: Array<any>) => {
  if (data.length === 0) throw new Error('Se necesita al menos una fila de datos');
  const planned = plainArraysToObjects(JSON.parse(JSON.stringify(data[0])));
  const worksheet = wb.addWorksheet(name, sheetConfig);
  renderPyramidHeaders('A1', worksheet, planned, `${name.toUpperCase()} SHEET`);
  renderData(worksheet, data);
  applyStyles(worksheet);
  reset();
};

const renderPyramidHeaders = (cell: string, worksheet: ExcelJS.Worksheet, data: any, name: string) => {
  const keys = Object.keys(data);
  const baseKeysCount = countBaseKeys(data);

  write(worksheet, cell, name);
  worksheet.mergeCells(`${cell}:${move(cell, Movements.RIGTH, baseKeysCount - 1)}`);
  cell = moveOne(cell, Movements.BOTTOM);

  for (let i = 0; i < keys.length; i++) {
    if (isObject(data[keys[i]])) {
      cell = renderPyramidHeaders(cell, worksheet, data[keys[i]], keys[i]);
      worksheet.getCell(cell).style = {
        font: {
          bold: true,
        },
      };
    } else {
      write(worksheet, cell, keys[i]);
      if (i + 1 !== keys.length) cell = moveOne(cell, Movements.RIGTH);
    }
  }

  if (lastAction === Actions.WRITE) cell = moveOne(cell, Movements.RIGTH);
  cell = moveOne(cell, Movements.TOP);
  headerFinalRow = lastRow;
  return cell;
};

const renderData = (worksheet: ExcelJS.Worksheet, data: Array<any>) => {
  const rows = data.map((row) => parsePlainArray(row));
  let cell = `A${headerFinalRow + 1}`;
  rows.forEach((row) => (cell = renderRow(cell, worksheet, row)));
};

const renderRow = (cell: string, worksheet: ExcelJS.Worksheet, row: Array<any>): string => {
  const baseRow = getNumber(cell);
  let high = 0;

  row.forEach((column) => {
    if (Array.isArray(column)) {
      let highAux = 0;
      worksheet.getCell(cell).style = {
        ...worksheet.getCell(cell).style,
        border: borderTopStyle,
      };
      column.forEach((cellValue) => {
        write(worksheet, cell, String(cellValue));
        cell = moveOne(cell, Movements.BOTTOM);
        highAux++;
      });
      if (highAux > high) high = highAux;
      cell = moveOne(`${getLetter(cell)}${baseRow}`, Movements.RIGTH);
    } else {
      write(worksheet, cell, String(column));
      worksheet.getCell(cell).style = {
        ...worksheet.getCell(cell).style,
        border: borderTopStyle,
      };
      cell = moveOne(cell, Movements.RIGTH);
    }
  });

  return `A${baseRow + high}`;
};

const parsePlainArray = (data: any) => {
  if (typeof data !== 'object') return data;

  const newArray: any[] = [];
  for (const key in data) {
    if (typeof data[key] == 'object' && data[key]) {
      if (Array.isArray(data[key]) && data[key].length > 0) {
        const allObjects = data[key].filter((e: any) => typeof e == 'object' && !Array.isArray(e));
        if (allObjects.length == data[key].length) {
          if (allObjects.length == 1) newArray.push(...parsePlainArray(data[key]));
          else {
            const names = Object.keys(allObjects[0]);
            const aux: any[] = [];
            for (let index = 0; index < names.length; index++) {
              const pos = names[index];
              const aux2 = data[key].map((e: any) => {
                if (typeof e !== 'object') return e[pos];
                return parsePlainArray(e[pos]);
              });
              aux.push(aux2);
            }
            newArray.push(...aux);
          }
        } else {
          newArray.push(
            data[key].map((e: any) => {
              if (typeof e !== 'object') return e;
              return parsePlainArray(e);
            }),
          );
        }
      } else if (Array.isArray(data[key])) newArray.push([]);
      else newArray.push(...parsePlainArray(data[key]));
    } else newArray.push(data[key]);
  }
  return newArray;
};

const isObject = (data: any) => data && typeof data === 'object' && !Array.isArray(data);

const countBaseKeys = (object: any) => {
  let count = 0;

  for (const key in object) {
    if (typeof object[key] === 'object') {
      count += countBaseKeys(object[key]);
    } else {
      count++;
    }
  }

  return count;
};

const plainArraysToObjects = (object: any) => {
  for (const key in object) {
    if (Array.isArray(object[key])) {
      if (object[key].length > 0) {
        object[key] = object[key][0];
      } else {
        object[key] = '';
      }
    } else if (typeof object[key] === 'object') {
      plainArraysToObjects(object[key]);
    }
  }
  return object;
};

const moveOne = (cell: string, movement: Movements) => {
  lastAction = Actions.MOVE;
  const number = getNumber(cell);
  if (movement === Movements.BOTTOM && number + 1 > lastRow) lastRow = number + 1;
  return move(cell, movement);
};

const move = (cell: string, movement: Movements, skip = 1) => {
  const number = getNumber(cell);
  const letter = getLetter(cell);
  switch (movement) {
    case Movements.RIGTH:
      return `${nextLetter(letter, skip)}${number}`;
    case Movements.LEFT:
      return `${previousLetter(letter, skip)}${number}`;
    case Movements.TOP:
      return `${letter}${number - skip}`;
    case Movements.BOTTOM:
      return `${letter}${number + skip}`;
  }
};

const write = (worksheet: ExcelJS.Worksheet, cell: string, value: any) => {
  lastAction = Actions.WRITE;
  worksheet.getCell(cell).value = value;
};

const nextLetter = (letter: string, skip: number = 1): string => {
  let result = '';
  let carry = skip;

  for (let i = letter.length - 1; i >= 0; i--) {
    const char = letter[i];
    const charCode = char.charCodeAt(0) - 65;
    const newCharCode = (charCode + carry) % 26;
    carry = Math.floor((charCode + carry) / 26);

    result = String.fromCharCode(newCharCode + 65) + result;
  }

  if (carry > 0) {
    result = String.fromCharCode(carry + 64) + result;
  }

  return result;
};

const previousLetter = (letter: string, skip: number = 1): string => {
  let result = '';
  let borrow = skip;

  for (let i = letter.length - 1; i >= 0; i--) {
    const char = letter[i];
    const charCode = char.charCodeAt(0) - 65;

    if (charCode >= borrow) {
      result = String.fromCharCode(charCode - borrow + 65) + result;
      borrow = 0;
    } else {
      result = String.fromCharCode(charCode + 26 - borrow + 65) + result;
      borrow = 1;
    }
  }

  return result;
};

const applyStyles = (worksheet: ExcelJS.Worksheet) => {
  worksheet.columns.forEach((column) => {
    if (!column.eachCell) return;
    let maxLength = 0;
    let lastCellWithValue: null | string = null;
    let toMergeCount = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      if (Number(cell.row) <= headerFinalRow) {
        if (cell.value && cell.value.toString().length !== 0) {
          if (lastCellWithValue && toMergeCount > 1)
            worksheet.mergeCells(`${lastCellWithValue}:${move(cell.address, Movements.BOTTOM, toMergeCount)}`);

          lastCellWithValue = cell.address;
          toMergeCount = 0;
        } else toMergeCount++;
        cell.style = {
          ...cell.style,
          font: {
            bold: true,
          },
        };
      }
      cell.style = {
        ...cell.style,
        alignment: {
          horizontal: 'center',
          vertical: 'middle',
        },
      };
      const columnLength = cell.value ? cell.value.toString().length : 10;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    if (toMergeCount > 0 && lastCellWithValue)
      worksheet.mergeCells(`${lastCellWithValue}:${move(lastCellWithValue, Movements.BOTTOM, toMergeCount)}`);
    column.width = maxLength < 10 ? 10 : maxLength;
  });
};

const getNumber = (cell: string) => parseInt(cell.substring(cell.search(/\d/)));

const getLetter = (cell: string) => cell.substring(0, cell.search(/\d/));

const reset = () => {
  lastAction = Actions.WRITE;
  lastRow = 0;
  headerFinalRow = 0;
};

