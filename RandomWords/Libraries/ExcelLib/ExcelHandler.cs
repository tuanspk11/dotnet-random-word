using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RandomWords.Libraries
{
    internal class ExcelHandler
    {
        #region config
        public SpreadsheetDocument? CurrentDocument { get; private set; }

        public string? CurrentFilename { get; private set; }

        public WorksheetPart? CurrentSheet { get; private set; }

        public Stylesheet? CurrentStyleSheet { get; set; }
        #endregion

        public bool OpenDocument(string fileName, string sheetName)
        {
            if (OpenDocument(fileName))
            {
                return SetCurrentSheet(sheetName);
            }
            return false;
        }

        public bool OpenDocument(string fileName)
        {
            return OpenDocument(fileName, true);
        }

        public bool OpenDocument(string fileName, bool editable)
        {
            try
            {
                if (!File.Exists(fileName))
                {
                    return false;
                }
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, editable);
                CurrentDocument = spreadSheet;
                CurrentFilename = fileName;
                CurrentStyleSheet = spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
            }
            catch
            {
                return false;
            }
            return true;
        }

        public bool SetCurrentSheet(string sheetName)
        {
            if (CurrentDocument == null)
            {
                return false;
            }
            IEnumerable<Sheet> sheets = CurrentDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return false;
            }
            string relationshipId = sheets.First().Id.Value;
            CurrentSheet = (WorksheetPart)CurrentDocument.WorkbookPart.GetPartById(relationshipId);

            return true;
        }

        public static string GetCellValue(Cell cell)
        {
            if (cell == null)
            {
                return null;
            }
            if (cell.DataType == null)
            {
                return cell.InnerText.Trim();
            }

            string value = cell.InnerText;
            switch (cell.DataType.Value)
            {
                //String
                case CellValues.SharedString:
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent && String.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        parent = parent.Parent;
                    }
                    if (String.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    if (sstPart != null && sstPart.SharedStringTable != null)
                    {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).First().InnerText.Trim();
                    }
                    break;

                //Boolean
                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }

            return value;
        }
    }
}
