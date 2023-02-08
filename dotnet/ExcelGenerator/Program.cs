namespace ExcelGenerator {
    class Program
    {
        static void Main(string[] args)
        {
            IParser? parser = new GeneratorParser();
            while (parser != null)
                parser = parser.Parse(new FunctionCall(Console.ReadLine()));
        }
    }
}