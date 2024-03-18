const MAX_ROW = 1048576;

function GetColumnRange(sheetName: string, startCell: string) {
  const ws = ThisWorkbook.Sheets.Item(sheetName);
  if (!ws) return;
  const start = ws.Range(startCell);
  const end = start.End(xlDown);

  if (end.Row === MAX_ROW) return [];
  return ws.Range(start, end);
}

function GetColumnRangeValues(sheetName: string, startCell: string) {
  const ws = ThisWorkbook.Sheets.Item(sheetName);
  if (!ws) return;
  const start = ws.Range(startCell);
  const end = start.End(xlDown);

  if (end.Row !== MAX_ROW) return ws.Range(start, end).Value();

  const cellValue = start.Value();
  return cellValue ? [[cellValue]] : [];
}

function GetColumnEndingRowNumber(sheetName: string, startCell: string) {
  const ws = ThisWorkbook.Sheets.Item(sheetName);
  if (!ws) return;

  const row = ws.Range(startCell).End(xlDown).Row;
  if (row === MAX_ROW) return ws.Range(startCell).Row;
  return row;
}

function GetAccumulatedStock(sheetName: string, nameCell: string, valueCell: string) {
  const keyRange = GetColumnRangeValues(sheetName, nameCell) as RowValue[];
  const numberRange = GetColumnRangeValues(sheetName, valueCell) as RowValue[];
  const resultValues = [];

  const countMap = new Map<string, number>();
  keyRange.forEach((row: string[]) => {
    countMap.set(row[0], 0);
  });

  for (let i = 0; i < numberRange.length; i++) {
    const key = keyRange[i][0] as string;
    const num = parseInt(numberRange[i][0] as string, 10);
    countMap.set(key, countMap.get(key) + num);
    resultValues.push([countMap.get(key)]);
  }

  return resultValues;
}

function LookUpStockName(
  registrySheetName: string,
  numberCellStart: string,
  nameCellStart: string,
  matchingSheetName: string,
  matchingCellStart: string
) {
  const numberRange = GetColumnRangeValues(registrySheetName, numberCellStart) as RowValue[];
  const nameRange = GetColumnRangeValues(registrySheetName, nameCellStart) as RowValue[];
  const stockMap = new Map<string, string>();

  for (let i = 0; i < numberRange.length; i++) {
    const itemNumber = numberRange[i][0] as string;
    const itemName = nameRange[i][0] as string;
    stockMap.set(itemNumber, itemName);
  }

  const matchingRange = GetColumnRangeValues(matchingSheetName, matchingCellStart) as RowValue[];
  const resultValues = [];
  for (let i = 0; i < matchingRange.length; i++) {
    const matching = matchingRange[i][0] as string;
    const name = stockMap.get(matching) || '无匹配项';
    resultValues.push([name]);
  }
  return resultValues;
}

function Application_WorkbookAfterSave(Wb, Success) {
  Application.CalculateFull();
}
