declare var xlDown: number;

declare type CellValue = string | number | undefined;

declare type RowValue = CellValue[];

declare type RangeValue = RowValue[] | CellValue | RowValue;

declare type RangeObject = {
  Value(): RangeValue;
  End(v1: number): RangeObject;
  Row: number;
};

declare type WorkbookSheet = {
  Range(v1: string | RangeObject, v2?: string | RangeObject): RangeObject;
};

declare type Workbook = {
  Range(v1: string | RangeObject, v2?: string | RangeObject): RangeObject;
};

declare var Application: {
  CalculateFull(): void;
};

declare var ThisWorkbook: {
  Sheets: {
    Item(name: string): WorkbookSheet;
  };
};
