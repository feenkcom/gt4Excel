using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator
{
    internal class DocumentParser : IParser
    {
        public DocumentParser(SpreadsheetDocument doc) {
            Document = doc;
            WorkbookPart = Document.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();
            Sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
            SharedStringTablePart = WorkbookPart.AddNewPart<SharedStringTablePart>();
        }
        public SpreadsheetDocument Document { get; set; }
        public WorkbookPart WorkbookPart { get; set; }
        public Sheets Sheets { get; set; }
        public SharedStringTablePart SharedStringTablePart { get; set; }

        public IParser? Parse(FunctionCall input)
        {
            switch (input.Function)
            {
                case "Save":
                    Document.Save();
                    Document.Close();
                    return new GeneratorParser();
                case "CreateSheet":
                    return new SheetParser(this, input.Arguments[0]);
            }
            return this;
        }
    }
}
