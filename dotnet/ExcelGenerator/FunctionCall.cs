namespace ExcelGenerator
{
    internal class FunctionCall
    {
        public FunctionCall(string? text) {
            int index;
            if (text == null || (index = text.IndexOf("(")) == -1)
            {
                Function = (text ?? "").Trim();
                Arguments = new string[0];
            } else
            {
                Function = text.Substring(0, index).Trim();
                var args = text.Substring(index + 1, Math.Max(text.LastIndexOf(")") - index - 1, 0)).Trim();
                if (args.Length == 0)
                    Arguments = new string[0];
                else
                    Arguments = ParseArgs(args);
            }
        }

        private string[] ParseArgs(string args)
        {
            var result = new List<string>();
            var index = 0;
            var startIndex = 0;
            while (index < args.Length)
            {
                while (index < args.Length && char.IsSeparator(args[index]))
                    index++;
                if (index >= args.Length)
                    break;
                startIndex = index;
                if (args[index] == '\"')
                {
                    while (index < args.Length && args[index] == '\"')
                    {
                        do
                        {
                            index++;
                        } while (index < args.Length && args[index] != '\"');
                        if (index < args.Length)
                            index++;
                    }
                    result.Add(ObjectFrom(args.Substring(startIndex, index - startIndex)));
                } else
                {
                    index = args.IndexOf(',', startIndex);
                    if (index < 0)
                        index = args.Length;
                    result.Add(args.Substring(startIndex, index - startIndex));
                }
                while (index < args.Length && char.IsSeparator(args[index]))
                    index++;
                if (index < args.Length && args[index] == ',')
                    startIndex = ++index;
                else
                    break;
            }
            return result.ToArray();
        }

        private string ObjectFrom(string value)
        {
            if (value.Length == 0)
                return "";
            if (value[0] == '\"')
                return value.Substring(1, value.Length - 2).Replace("\"\"", "\"").Replace("\\r", "\r").Replace("\\n", "\n");
            return value;
        }

        public string Function { get; set; }
        public string[] Arguments { get; set; }
    }
}
