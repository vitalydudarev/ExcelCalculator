using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetCalculator.WorkbookTypes
{
    public class Worksheet
    {
        public string Name { get; }
        public Cell[] Cells { get; }

        private readonly Dictionary<string, Cell> _cells;

        public Worksheet(string name, Cell[] cells)
        {
            Name = name;
            Cells = cells;

            _cells = cells.ToDictionary(a => a.CellReference, b => b);
        }

        public Cell GetCell(string cellReference)
        {
            return _cells[cellReference];
        }
    }
}