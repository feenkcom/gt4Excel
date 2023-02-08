namespace ExcelGenerator
{
    internal interface IParser
    {
        IParser? Parse(FunctionCall input);
    }
}
