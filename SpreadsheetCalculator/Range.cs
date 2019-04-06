using SpreadsheetCalculator.WorkbookTypes;

namespace SpreadsheetCalculator
{
    public class Range
    {
        public Cell[] Cells { get; }

        public Range(Cell[] cells)
        {
            Cells = cells;
        }
    }
}