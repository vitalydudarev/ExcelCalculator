using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace SpreadsheetCalculator
{
    class Program
    {
        static void Main(string[] args)
        {
            var tempPath = Path.GetTempPath();
            var fileName = @"test.xlsm";
            var shortFileName = Path.GetFileNameWithoutExtension(fileName);
            var workbookFolder = Path.Combine(tempPath, shortFileName);
            if (Directory.Exists(workbookFolder))
                Directory.Delete(workbookFolder, true);

            System.IO.Compression.ZipFile.ExtractToDirectory(fileName, workbookFolder);


            var workbookReader = new WorkbookReader();
            var workbook = workbookReader.Read(workbookFolder);

            workbook.GetDefinedNameValues(workbook.DefinedNames[0].Name);
            
            Console.WriteLine("Hello World!");
        }
    }
    
    public class Worksheet
    {
        public string Name { get; }
        public Cell[] Cells { get; }

        public Worksheet(string name, Cell[] cells)
        {
            Name = name;
            Cells = cells;
        }
    }

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

    public interface IReference
    {
    }

    [DebuggerDisplay("{" + nameof(Reference) + "}")]
    public class Cell : IReference
    {
        public string Reference { get; }
        public CellType CellType { get; }
        public string CellValue { get; set; }
        public CellFormula CellFormula { get; }
        public bool HasFormula => CellFormula != null;
        public bool HasValue => !string.IsNullOrEmpty(CellValue);

        private readonly string[] _sharedStrings;

        public Cell(string reference, CellType cellType, string cellValue, CellFormula cellFormula,
            string[] sharedStrings)
        {
            Reference = reference;
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

    public class Range : IReference
    {
        public Cell[] Cells { get; set; }
    }

    public interface IValue
    {
    }

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
    
    public class Workbook
    {
        public Worksheet[] Sheets { get; }
        public DefinedName[] DefinedNames { get; }
        
        private CalcChainCell[] _calcChainCells;
        private string[] _sharedString;
        private readonly Dictionary<string, string> _definedNames;

//        public Workbook(Worksheet[] worksheets, DefinedName[] definedNames)
//        {
//            Sheets = worksheets;
//            DefinedNames = definedNames;
//        }

        public Workbook(Worksheet[] worksheets, DefinedName[] definedNames, CalcChainCell[] calcChainCells,
            string[] sharedStrings)
        {
            Sheets = worksheets;
            DefinedNames = definedNames.Where(a => !a.Hidden).ToArray();
            
            _calcChainCells = calcChainCells;
            _sharedString = sharedStrings;
            _definedNames = definedNames.Where(a => !a.Hidden).ToDictionary(a => a.Name, b => b.Reference);
        }

        public void Calculate()
        {
        }

        public XlValue GetDefinedNameValues(string name)
        {
            var reference = _definedNames[name];
            ParseReference(reference);

            throw new NotImplementedException();
        }

        private void ParseReference(string reference)
        {
            var sheetRange = reference.Split('!');
            if (sheetRange.Length == 2)
            {
                var sheetName = sheetRange[0].Trim('\'');
                var range = sheetRange[1].Replace("$", "").Split(':');
                if (range.Length == 2)
                {
                    var rangeFrom = range[0];
                    var rangeTo = range[1];
                }
            }
        }
    }

    public class Reference
    {
        public string SheetName { get; }
        public string[] Cells { get; }
    }

    public class XlValue
    {
        public object Value { get; }
        public CellType CellType { get; }
    }

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

    public enum CellFormulaType
    {
        Array,
        TableFormula,
        Normal,
        SharedFormula
    }

    // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CellType_topic_ID0E6NEFB.html
    public enum CellType
    {
        Boolean,
        Error,
        InlineStr,
        Number,
        SharedString,
        String
    }
}