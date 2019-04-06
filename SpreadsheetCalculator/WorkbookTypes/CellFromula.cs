namespace SpreadsheetCalculator.WorkbookTypes
{
    public class CellFormula
    {
        public bool CalculateCell { get; }
        public bool AlwaysCalculateArray { get; }
        public string RangeOfCells { get; }
        public string Formula { get; }
        public CellFormulaType FormulaType { get; }

        public CellFormula(bool calculateCell, bool alwaysCalculateArray, string rangeOfCells, string formula,
            CellFormulaType formulaType)
        {
            CalculateCell = calculateCell;
            AlwaysCalculateArray = alwaysCalculateArray;
            RangeOfCells = rangeOfCells;
            Formula = formula;
            FormulaType = formulaType;
        }
    }
}