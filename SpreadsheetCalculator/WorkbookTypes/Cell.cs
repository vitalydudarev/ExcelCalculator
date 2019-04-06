using System;
using System.Diagnostics;

namespace SpreadsheetCalculator.WorkbookTypes
{
    [DebuggerDisplay("{" + nameof(CellReference) + "}")]
    public class Cell
    {
        public string CellReference { get; }
        public CellType CellType { get; }
        public string CellValue { get; set; }
        public CellFormula CellFormula { get; }
        public bool HasFormula => CellFormula != null;
        public bool HasValue => !string.IsNullOrEmpty(CellValue);

        private readonly string[] _sharedStrings;

        public Cell(string reference, CellType cellType, string cellValue, CellFormula cellFormula,
            string[] sharedStrings)
        {
            CellReference = reference;
            CellType = cellType;
            CellValue = cellValue;
            CellFormula = cellFormula;
            _sharedStrings = sharedStrings;
        }

        // TODO: return IValue
        public object GetValue()
        {
            if (CellType == CellType.SharedString)
                return _sharedStrings[int.Parse(CellValue)];

            if (CellType == CellType.String)
                return CellValue;

//            if (CellDataType == )

            throw new NotImplementedException();
        }

        public void SetValue(string value)
        {
            CellValue = value;
        }
    }
}