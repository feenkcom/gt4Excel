using ClosedXML.Excel;

namespace ExcelGenerator
{
    internal class GeneratorParser : IParser
    {
        public IParser? Parse(FunctionCall input)
        {
            switch (input.Function)
            {
                case "CreateDocument":
                    var workbook = new XLWorkbook();
                    return new DocumentParser(workbook, input.Arguments[0]);
                case "Quit":
                    return null;
            }
            return this;
        }
    }
}
