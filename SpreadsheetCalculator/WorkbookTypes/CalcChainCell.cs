namespace SpreadsheetCalculator.WorkbookTypes
{
    public class CalcChainCell
    {
        public string SheetName { get; }
        public string CellReference { get; }
        public bool IsArrayFormula { get; }
        public bool IsNewLevelDependency { get; }
        public bool IsChildChain { get; }
        public bool IsNewThread { get; }

        public CalcChainCell(string sheetName, string cellReference, bool isArrayFormula, bool isNewLevelDependency,
            bool isChildChain, bool isNewThread)
        {
            SheetName = sheetName;
            CellReference = cellReference;
            IsArrayFormula = isArrayFormula;
            IsNewLevelDependency = isNewLevelDependency;
            IsChildChain = isChildChain;
            IsNewThread = isNewThread;
        }
    }
}