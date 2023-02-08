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
                    Arguments = args.Split(",").Select(each => ObjectFrom(each.Trim())).ToArray();
            }
        }

        private string ObjectFrom(string value)
        {
            if (value.Length == 0)
                return "";
            if (value[0] == '\"')
                return value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
            return value;
        }

        public string Function { get; set; }
        public string[] Arguments { get; set; }
    }
}
