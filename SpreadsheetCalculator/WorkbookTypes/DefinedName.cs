namespace SpreadsheetCalculator.WorkbookTypes
{
    public class DefinedName
    {
        public string Name { get; }
        public string Reference { get; }
        public bool Hidden { get; }

        public DefinedName(string name, string reference, bool hidden)
        {
            Name = name;
            Reference = reference;
            Hidden = hidden;
        }
    }
}