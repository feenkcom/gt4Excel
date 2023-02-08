using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator
{
    internal class SheetParser : IParser
    {
        public SheetParser(DocumentParser outerParser, string name)
        {
            DocumentParser = outerParser;
            WorksheetPart = DocumentParser.WorkbookPart.AddNewPart<WorksheetPart>();
            WorksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheet = new Sheet()
            {
                Id = DocumentParser.WorkbookPart.GetIdOfPart(WorksheetPart),
                SheetId = (uint)DocumentParser.Sheets.ChildElements.Count + 1,
                Name = name
            };
            DocumentParser.Sheets.Append(sheet);
            Sheet = sheet;
        }

        public DocumentParser DocumentParser { get; }
        public WorksheetPart WorksheetPart { get; }

        public Sheet Sheet { get; }

        public IParser? Parse(FunctionCall input)
        {
            Cell cell;
            switch (input.Function)
            {
                case "SetText":
                    var index = InsertSharedString(input.Arguments[0]);
                    cell = GetCell(input.Arguments[1], uint.Parse(input.Arguments[2]));
                    cell.CellValue = new CellValue(index);
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                case "SetNumber":
                    var value = decimal.Parse(input.Arguments[0]);
                    cell = GetCell(input.Arguments[1], uint.Parse(input.Arguments[2]));
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case "SetFormula":
                    cell = GetCell(input.Arguments[1], uint.Parse(input.Arguments[2]));
                    cell.CellFormula = new CellFormula(input.Arguments[0]);
                    break;
                case "EndSheet":
                    WorksheetPart.Worksheet.Save();
                    return DocumentParser;
            }
            return this;
        }

        private int InsertSharedString(string text)
        {
            if (DocumentParser.SharedStringTablePart.SharedStringTable == null)
                DocumentParser.SharedStringTablePart.SharedStringTable = new SharedStringTable();
            int i = 0;
            foreach (var item in DocumentParser.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                    return i;
                i++;
            }

            DocumentParser.SharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            DocumentParser.SharedStringTablePart.SharedStringTable.Save();

            return i;
        }

        private Cell GetCell(string columnName, uint rowIndex)
        {
            var worksheet = WorksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData!.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
 
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell? refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell!.CellReference!.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                var newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                return newCell;
            }
        }
    }
}
