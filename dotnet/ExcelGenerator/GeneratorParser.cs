using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace ExcelGenerator
{
    internal class GeneratorParser : IParser
    {
        public IParser? Parse(FunctionCall input)
        {
            switch (input.Function)
            {
                case "CreateDocument":
                    var spreadsheetDocument = SpreadsheetDocument.Create(input.Arguments[0], SpreadsheetDocumentType.Workbook);
                    return new DocumentParser(spreadsheetDocument);
                case "Quit":
                    return null;
            }
            return this;
        }
    }
}
