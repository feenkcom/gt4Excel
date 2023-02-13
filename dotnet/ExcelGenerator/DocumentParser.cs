using ClosedXML.Excel;

namespace ExcelGenerator
{
    internal class DocumentParser : IParser
    {
        public DocumentParser(XLWorkbook doc, string filename) {
            Workbook = doc;
            Filename = filename;
        }
        public XLWorkbook Workbook { get; }
        public string Filename;

        public IParser? Parse(FunctionCall input)
        {
            switch (input.Function)
            {
                case "Save":
                    Workbook.SaveAs(Filename);
                    Workbook.Dispose();
                    return new GeneratorParser();
                case "CreateSheet":
                    return new SheetParser(this, input.Arguments[0]);
            }
            return this;
        }
    }
}
