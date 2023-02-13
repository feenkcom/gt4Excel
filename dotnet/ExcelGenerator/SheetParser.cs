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
            switch (input.Function)
            {
                case "SetText":
                    Worksheet.Cell(input.Arguments[0]).Value = input.Arguments[1];
                    break;
                case "SetNumber":
                    Worksheet.Cell(input.Arguments[0]).Value = decimal.Parse(input.Arguments[1]);
                    break;
                case "SetFormula":
                    Worksheet.Cell(input.Arguments[0]).FormulaA1 = input.Arguments[1];
                    break;
                case "Bold":
                    Worksheet.Cell(input.Arguments[0]).Style.Font.Bold = true;
                    break;
                case "AdjustWidthToContents":
                    if (input.Arguments.Length == 0)
                        Worksheet.Columns().AdjustToContents();
                    else
                        foreach (var each in input.Arguments)
                            Worksheet.Columns(each).AdjustToContents();
                    break;
                case "AdjustHeightToContents":
                    if (input.Arguments.Length == 0)
                        Worksheet.Rows().AdjustToContents();
                    else
                        foreach (var each in input.Arguments)
                            Worksheet.Rows(each).AdjustToContents();
                    break;
                case "EndSheet":
                    return DocumentParser;
            }
            return this;
        }
    }
}
