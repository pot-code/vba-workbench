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

  if (end.Row === MAX_ROW) return [[ws.Range(startCell).Value()]];
  return ws.Range(start, end).Value();
}

function GetColumnEndingRowNumber(sheetName: string, startCell: string) {
  const ws = ThisWorkbook.Sheets.Item(sheetName);
  if (!ws) return;

  const row = ws.Range(startCell).End(xlDown).Row;
  if (row === MAX_ROW) return ws.Range(startCell).Row;
  return row;
}

function GetSumOfProductTillNow(sheetName: string) {
  const keyRange = GetColumnRangeValues(sheetName, 'B5') as RowValue[];
  const numberRange = GetColumnRangeValues(sheetName, 'F5') as RowValue[];
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

function Application_WorkbookAfterSave(Wb, Success) {
  Application.CalculateFull();
}
