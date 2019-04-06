namespace SpreadsheetCalculator
{
    public class Reference
    {
        public string SheetName { get; }
        public Range Range { get; }

        public Reference(string sheetName, Range range)
        {
            SheetName = sheetName;
            Range = range;
        }
    }
}