using ClosedXML.Excel;

namespace ExcelGenerator
{
    internal class SheetParser : IParser
    {
        public SheetParser(DocumentParser outerParser, string name)
        {
            DocumentParser = outerParser;
            Worksheet = DocumentParser.Workbook.AddWorksheet(name);
        }

        public DocumentParser DocumentParser { get; }
        public IXLWorksheet Worksheet { get; }

        public IParser? Parse(FunctionCall input)
        {
            IXLCell cell;
            switch (input.Function)
            {
                case "SetText":
                    cell = Worksheet.Cell(input.Arguments[1] + input.Arguments[2]);
                    cell.Value = input.Arguments[0];
                    break;
                case "SetNumber":
                    cell = Worksheet.Cell(input.Arguments[1] + input.Arguments[2]);
                    cell.Value = decimal.Parse(input.Arguments[0]);
                    break;
                case "SetFormula":
                    cell = Worksheet.Cell(input.Arguments[1] + input.Arguments[2]);
                    cell.FormulaA1 = input.Arguments[0];
                    break;
                case "EndSheet":
                    return DocumentParser;
            }
            return this;
        }
    }
}
