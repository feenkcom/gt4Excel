namespace ExcelGenerator {
    class Program
    {
        static void Main(string[] args)
        {
            Console.InputEncoding = System.Text.Encoding.UTF8;
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
                    Console.Error.WriteLine(ex);
                    Console.Error.Write("ERROR: ");
                    Console.Error.WriteLine(text);
                    Environment.Exit(-1);
                    break;
                }
            }
        }
    }
}