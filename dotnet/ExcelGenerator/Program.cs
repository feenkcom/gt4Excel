namespace ExcelGenerator {
    class Program
    {
        static void Main(string[] args)
        {
            IParser? parser = new GeneratorParser();
            while (parser != null)
            {
                string? text = Console.ReadLine();
                try
                {
                    parser = parser.Parse(new FunctionCall(text));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    Console.Write("ERROR: ");
                    Console.WriteLine(text);
                    break;
                }
            }
        }
    }
}