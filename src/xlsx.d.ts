declare module 'xlsx' {
  export interface WorkBook {
    SheetNames: string[];
    Sheets: { [sheet: string]: WorkSheet };
  }

  export interface WorkSheet {
    [cell: string]: CellObject | any;
    '!merges'?: Range[];
    '!ref'?: string;
  }

  export interface CellObject {
    v?: any; // value
    t?: string; // type
    s?: CellStyle; // style
  }

  export interface CellStyle {
    font?: {
      name?: string;
      sz?: number;
      color?: { rgb: string };
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
    };
    fill?: {
      patternType?: string;
      fgColor?: { rgb: string };
      bgColor?: { rgb: string };
    };
    alignment?: {
      horizontal?: string;
      vertical?: string;
      wrapText?: boolean;
    };
  }

  export interface Range {
    s: { r: number; c: number }; // start row,col
    e: { r: number; c: number }; // end row,col
  }

  export interface JSON2SheetOpts {
    header?: string[];
    skipHeader?: boolean;
    origin?: number | string;
  }

  export const utils: {
    json_to_sheet(data: any[], opts?: JSON2SheetOpts): WorkSheet;
    aoa_to_sheet(data: any[][], opts?: any): WorkSheet;
    sheet_add_aoa(
      worksheet: WorkSheet,
      data: any[][],
      opts?: { origin?: number | string }
    ): void;
    sheet_add_json(
      worksheet: WorkSheet,
      data: any[],
      opts?: JSON2SheetOpts
    ): void;
  };

  export function write(
    workbook: WorkBook,
    options: { bookType: string; type: string; cellStyles?: boolean }
  ): any;
}
