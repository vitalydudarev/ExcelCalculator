using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetCalculator.WorkbookTypes
{
    public class Workbook
    {
        public Worksheet[] Sheets { get; }
        public DefinedName[] DefinedNames { get; }
        
        private CalcChainCell[] _calcChainCells;
        private string[] _sharedString;
        private readonly Dictionary<string, string> _definedNames;
        private readonly Dictionary<string, Worksheet> _sheets;
        private Dictionary<string, int> _columns = GetWorksheetColumns();

        public Workbook(Worksheet[] worksheets, DefinedName[] definedNames, CalcChainCell[] calcChainCells,
            string[] sharedStrings)
        {
            Sheets = worksheets;
            DefinedNames = definedNames.Where(a => !a.Hidden).ToArray();
            
            _calcChainCells = calcChainCells;
            _sharedString = sharedStrings;
            _definedNames = definedNames.Where(a => !a.Hidden).ToDictionary(a => a.Name, b => b.Reference);
            _sheets = worksheets.ToDictionary(a => a.Name, b => b);
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

        private Range GetRange(string sheetName, string from, string to = null)
        {
            if (to == null)
            {
                var cell = _sheets[sheetName].GetCell(from);
                return new Range(new Cell[1] { cell });
            }

            var columns = GetRangeColumns(from, to);
            var cells = new Cell[columns.Length];
            //var sheets = _sheets[sheetName].GetCell("ddddd");

            throw new Exception();
        }

        private string[] GetRangeColumns(string from, string to)
        {
            var fromIndex = _columns[from];
            var toIndex = _columns[to];

            int columnCount = toIndex - fromIndex + 1;
            string[] columnNames = new string[columnCount];

            for (int i = 0; i < columnCount; i++)
                columnNames[i] = _columns.ElementAt(i).Key;

            return columnNames;
        }

        private static Dictionary<string, int> GetWorksheetColumns()
        {
            var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var result = new Dictionary<string, int>();

            for (int i = 0; i < Math.Pow(2, 14); i++)
            {
                var column = GetString(i, alphabet);
                result.Add(column, i);
            }

            return result;
        }

        private static string GetString(int index, string alphabet)
        {
            int alphabetLength = alphabet.Length;

            int a = index % alphabetLength;
            int b = index / alphabetLength;

            string result = "";

            if (b > 0)
                result += GetString(b - 1, alphabet);

            return result + alphabet[a];
        }
    }
}