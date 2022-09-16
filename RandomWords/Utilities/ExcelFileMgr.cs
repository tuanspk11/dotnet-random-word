using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

using RandomWords.Libraries;

namespace RandomWords.Utilities
{
    internal class ExcelFileMgr
    {
        private static Dictionary<int, string>? dicCellReference = null;

        private static Dictionary<int, string> DicCellReference
        {
            get
            {
                if (dicCellReference == null)
                {
                    dicCellReference = new Dictionary<int, string>();
                    dicCellReference.Add(1, "A");
                    dicCellReference.Add(2, "B");
                    dicCellReference.Add(3, "C");
                    dicCellReference.Add(4, "D");
                    dicCellReference.Add(5, "E");
                    dicCellReference.Add(6, "F");
                    dicCellReference.Add(7, "G");
                    dicCellReference.Add(8, "H");
                    dicCellReference.Add(9, "I");
                    dicCellReference.Add(10, "J");
                    dicCellReference.Add(11, "K");
                    dicCellReference.Add(12, "L");
                    dicCellReference.Add(13, "M");
                    dicCellReference.Add(14, "N");
                    dicCellReference.Add(15, "O");
                    dicCellReference.Add(16, "P");
                    dicCellReference.Add(17, "Q");
                    dicCellReference.Add(18, "R");
                    dicCellReference.Add(19, "S");
                    dicCellReference.Add(20, "T");
                    dicCellReference.Add(21, "U");
                    dicCellReference.Add(22, "V");
                    dicCellReference.Add(23, "W");
                    dicCellReference.Add(24, "X");
                    dicCellReference.Add(25, "Y");
                    dicCellReference.Add(26, "Z");
                }
                return dicCellReference;
            }
        }

        public Dictionary<string, DataTable> ReadExcelFile(string filePath, List<string> lstDataSheets, out string message)
        {
            var result = new Dictionary<string, DataTable>();
            message = string.Empty;

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    foreach (Sheet sheet in sheets)
                    {
                        Worksheet worksheet = ((WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                        if (lstDataSheets.Contains(sheet.Name))
                        {
                            DataTable dtTemp = GetSheetDataTable(sheetData, 0, out message);
                            if (dtTemp != null)
                            {
                                result.Add(sheet.Name, dtTemp);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return result;
        }

        private DataTable? GetSheetDataTable(SheetData sheetData, int endCol, out string message)
        {
            var result = new DataTable();
            message = string.Empty;

            string value = string.Empty;
            bool isRow = true;
            string cellReference = string.Empty;

            IEnumerable<Row> rows = sheetData.Elements<Row>();
            foreach (Row row in rows)
            {
                if (row.RowIndex == 1)
                {
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        value = ExcelHandler.GetCellValue(c);
                        if (!string.IsNullOrEmpty(value))
                        {
                            result.Columns.Add(value);
                        }
                    }
                }
                else
                {
                    DataRow dr = result.NewRow();
                    int col = 0;
                    isRow = false;

                    foreach (Cell c in row.Elements<Cell>())
                    {
                        value = ExcelHandler.GetCellValue(c);
                        cellReference = makeCellReference(col) + row.RowIndex.ToString();

                        if (col == endCol && (string.IsNullOrEmpty(value) || c.CellReference != cellReference))
                        {
                            return null;
                        }

                        if (result.Columns.Count > col)
                        {
                            while (c.CellReference != cellReference && col >= endCol)
                            {
                                dr[result.Columns[col]] = "";
                                col += 1;
                                cellReference = makeCellReference(col) + row.RowIndex.ToString();
                            }
                            isRow = true;
                            dr[result.Columns[col]] = value;
                            col += 1;
                        }
                    }

                    if (isRow && col > endCol)
                    {
                        result.Rows.Add(dr);
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            return result;
        }

        private string makeCellReference(int col)
        {
            col = col + 1;
            string Reference = string.Empty;
            int header = col / 26;
            int child = col % 26;
            if (header == 0)
            {
                Reference = DicCellReference[child]; // A - Y 
            }
            else if (header == 1 && child == 0)
            {
                Reference = DicCellReference[26]; // Z
            }
            else if (header != 1 && child == 0)
            {
                Reference = DicCellReference[header - 1] + DicCellReference[26]; // AZ
            }
            else
            {
                Reference = DicCellReference[header] + DicCellReference[child]; // B? - 
            }
            return Reference;
        }
    }
}
