using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SpreadsheetCalculator.WorkbookTypes;

namespace SpreadsheetCalculator
{
    public class WorkbookReader
    {
        public Workbook Read(string fileName)
        {
            var sharedStringsFileName = Path.Combine(fileName, "xl\\sharedStrings.xml");
            var calcChainFileName = Path.Combine(fileName, "xl\\calcChain.xml");
            var workbookFileName = Path.Combine(fileName, "xl\\workbook.xml");
            var contentTypesFileName = Path.Combine(fileName, "[Content_Types].xml");
            var workbookRelsFileName = Path.Combine(fileName, "xl\\_rels\\workbook.xml.rels");

            var workbookInfo = ParseWorkbook(workbookFileName);
            var workbookRels = ParseWorkbookRels(workbookRelsFileName);
            var sharedStrings = ParseSharedStrings(sharedStringsFileName);
            var calcChainCells = ParseCalculationChain(calcChainFileName);

//            var worksheets = new List<Worksheet>();

//            var cells = new Dictionary<string, Cell>();
//            var cellols = new List<ExcelReader.Cell>();

            var worksheets = new List<WorkbookTypes.Worksheet>();
            
            foreach (var worksheetInfo in workbookInfo.WorksheetInfos)
            {
                var sheet = worksheetInfo.Name;
                var rId = worksheetInfo.Rid;
                var rel = workbookRels.FirstOrDefault(a => a.Id == rId);
                if (rel != null)
                {
                    var worksheetFileName = Path.Combine(fileName, "xl", rel.Target.Replace('/', '\\'));
                    var rows = ParseWorksheet(worksheetFileName);
                    
                    var cecells = new List<WorkbookTypes.Cell>();

                    foreach (var row in rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            WorkbookTypes.CellFormula formula = null;

                            if (cell.HasFormula)
                            {
                                var cellFormula = cell.CellFormula;
                                formula = new WorkbookTypes.CellFormula(cellFormula.CalculateCell,
                                    cellFormula.AlwaysCalculateArray, cellFormula.RangeOfCells,
                                    cellFormula.Formula, cellFormula.FormulaType);
                            }

                            var cell1 = new WorkbookTypes.Cell(cell.Reference, cell.CellDataType, cell.CellValue,
                                formula, sharedStrings);
                            cecells.Add(cell1);
                        }
                    }

                    var ws = new WorkbookTypes.Worksheet(sheet, cecells.ToArray());
                    worksheets.Add(ws);
                }
            }

            var mapping = workbookInfo.WorksheetInfos.ToDictionary(a => a.SheetId, b => b.Name);
            var calcCells = new List<WorkbookTypes.CalcChainCell>();
            
            foreach (var calcChainCell in calcChainCells)
            {
                var sheetName = mapping[calcChainCell.SheetId];

                var calcCell = new WorkbookTypes.CalcChainCell(sheetName, calcChainCell.CellReference,
                    calcChainCell.IsArrayFormula, calcChainCell.IsNewLevelDependency, calcChainCell.IsChildChain,
                    calcChainCell.IsNewThread);
                calcCells.Add(calcCell);
            }
            
            var defNames = workbookInfo.DefinedNames;
//            var ddd = cellols.Where(a => a.HasFormula && !string.IsNullOrEmpty(a.CellFormula.RangeOfCells))
//                .Select(a => a.CellFormula.RangeOfCells).ToArray();

            return new Workbook(worksheets.ToArray(), workbookInfo.DefinedNames, calcCells.ToArray(), sharedStrings);
        }

        private static IEnumerable<Row> ParseWorksheet(string fileName)
        {
            var xmlParser = new XmlParser(fileName);

            var rows = new List<Row>();
            var rowNodes = xmlParser.GetNodes("row");

            foreach (var rowNode in rowNodes)
            {
                var cells = new List<Cell>();

                foreach (var cNode in xmlParser.GetNodes(rowNode, "c"))
                {
                    CellFormula cellFormula = null;
                    string cellValue = null;

                    // take formula
                    var fNode = xmlParser.GetSingleNode(cNode, "f");
                    if (fNode != null)
                    {
                        var formula = fNode.InnerText;
                        var fNodeAttributes = xmlParser.GetAttributes(fNode);

                        cellFormula = new CellFormula(fNodeAttributes, formula);
                    }

                    var isNode = xmlParser.GetSingleNode(cNode, "is");
                    if (isNode != null)
                    {
                        throw new Exception("is tag is not supported.");
                    }

                    // take value
                    var vNode = xmlParser.GetSingleNode(cNode, "v");
                    if (vNode != null)
                        cellValue = vNode.InnerText;

                    var cNodeAttributes = xmlParser.GetAttributes(cNode);

                    var cell = new Cell(cNodeAttributes, cellFormula, cellValue);
                    cells.Add(cell);
                }

                var rowAttributes = xmlParser.GetAttributes(rowNode);
                var row = new Row(rowAttributes, cells.ToArray());
                rows.Add(row);
            }

            return rows;
        }

        private static Relationship[] ParseWorkbookRels(string fileName)
        {
            var xmlParser = new XmlParser(fileName);

            var relationships = new List<Relationship>();

            foreach (var relationshipNode in xmlParser.GetNodes("Relationship"))
            {
                var relationshipAttributes = xmlParser.GetAttributes(relationshipNode);
                var relationship = new Relationship(relationshipAttributes);

                relationships.Add(relationship);
            }

            return relationships.ToArray();
        }

        private static WorkbookInfo ParseWorkbook(string fileName)
        {
            var xmlParser = new XmlParser(fileName);

            var worksheetInfos = new List<WorksheetInfo>();
            var definedNames = new List<DefinedName>();
            
            foreach (var sheetNode in xmlParser.GetNodes("sheet"))
            {
                var sheetNodeAttributes = xmlParser.GetAttributes(sheetNode);
                var worksheetInfo = new WorksheetInfo(sheetNodeAttributes);

                worksheetInfos.Add(worksheetInfo);
            }
            
            foreach (var definedNameNode in xmlParser.GetNodes("definedName"))
            {
                var definedNameNodeAttributes = xmlParser.GetAttributes(definedNameNode);

                var name = definedNameNodeAttributes["name"];
                var hidden = definedNameNodeAttributes.TryGetValue("hidden", out var hidden2) && hidden2 == "1";
                var definedName = new DefinedName(name, definedNameNode.InnerText, hidden);

                definedNames.Add(definedName);
            }

            return new WorkbookInfo(worksheetInfos.ToArray(), definedNames.ToArray());
        }

        private static string[] ParseSharedStrings(string fileName)
        {
            var xmlParser = new XmlParser(fileName);

            var sharedStrings = new List<string>();
            
            foreach (var siNode in xmlParser.GetNodes("si"))
            {
                string value = "";

                foreach (var tNode in xmlParser.GetNodes(siNode, "t"))
                    value += tNode.InnerText;

                sharedStrings.Add(value);
            }

            return sharedStrings.ToArray();
        }

        private static CalcChainCell[] ParseCalculationChain(string fileName)
        {
            var xmlParser = new XmlParser(fileName);

            var calcChainCells = new List<CalcChainCell>();
            var cNodes = xmlParser.GetNodes("c");

            foreach (var cNode in cNodes)
            {
                var cNodeAttributes = xmlParser.GetAttributes(cNode);
                var calcChainCell = new CalcChainCell(cNodeAttributes);

                calcChainCells.Add(calcChainCell);
            }

            return calcChainCells.ToArray();
        }

        #region Private Classes

        private class WorkbookInfo
        {
            public WorksheetInfo[] WorksheetInfos { get; }
            public DefinedName[] DefinedNames { get; }

            public WorkbookInfo(WorksheetInfo[] worksheetInfos, DefinedName[] definedNames)
            {
                WorksheetInfos = worksheetInfos;
                DefinedNames = definedNames;
            }
        }

        private class Relationship
        {
            public string Id { get; }
            public string Type { get; }
            public string Target { get; }

            public Relationship(Dictionary<string, string> attributes)
            {
                foreach (var attribute in attributes)
                {
                    if (attribute.Key == "Id")
                        Id = attribute.Value;
                    if (attribute.Key == "Type")
                        Type = attribute.Value;
                    if (attribute.Key == "Target")
                        Target = attribute.Value;
                }
            }
        }

        private class WorksheetInfo
        {
            public string Name { get; }
            public int SheetId { get; }
            public string Rid { get; }

            public WorksheetInfo(Dictionary<string, string> attributes)
            {
                foreach (var attribute in attributes)
                {
                    var key = attribute.Key;
                    var value = attribute.Value;

                    if (key == "name")
                        Name = value;
                    if (key == "sheetId")
                        SheetId = int.Parse(value);
                    if (key == "r:id")
                        Rid = value;
                }
            }
        }
        
        private class Cell
        {
            public string Reference { get; }
            public CellType CellDataType { get; }
            public string CellValue { get; }
            public CellFormula CellFormula { get; }
            public bool HasFormula { get; }

            public Cell(Dictionary<string, string> attributes, CellFormula cellFormula, string cellValue)
            {
                foreach (var attribute in attributes)
                {
                    if (attribute.Key == "r")
                        Reference = attribute.Value;
                    if (attribute.Key == "t")
                        CellDataType = ParseCellDataType(attribute.Value);
                }

                CellFormula = cellFormula;
                CellValue = cellValue;
                HasFormula = cellFormula != null;
            }

            private static CellType ParseCellDataType(string s)
            {
                if (s == "b")
                    return CellType.Boolean;
                if (s == "e")
                    return CellType.Error;
                if (s == "inlineStr")
                    return CellType.InlineStr;
                if (s == "n")
                    return CellType.Number;
                if (s == "s")
                    return CellType.SharedString;
                if (s == "str")
                    return CellType.String;

                return CellType.Number;
            }
        }

        private class Row
        {
            public int RowIndex { get; }
            public Cell[] Cells { get; }

            public Row(Dictionary<string, string> rowAttributes, Cell[] cells)
            {
                foreach (var attribute in rowAttributes)
                {
                    if (attribute.Key == "r")
                        RowIndex = int.Parse(attribute.Value);
                }

                Cells = cells;
            }
        }

        private class Worksheet
        {
            public string Name { get; }
            public int SheetId { get; }
            public Row[] Rows { get; }

            public Worksheet(string name, int sheetId, Row[] rows)
            {
                Name = name;
                SheetId = sheetId;
                Rows = rows;
            }
        }

        private class CalcChainCell
        {
            public int SheetId { get; }
            public bool IsArrayFormula { get; }
            public bool IsNewLevelDependency { get; }
            public string CellReference { get; }
            public bool IsChildChain { get; }
            public bool IsNewThread { get; }

            public CalcChainCell(Dictionary<string, string> attributes)
            {
                foreach (var attribute in attributes)
                {
                    var attributeValue = attribute.Value;

                    switch (attribute.Key)
                    {
                        case "i":
                            SheetId = Convert.ToInt32(attributeValue);
                            break;
                        case "a":
                            IsArrayFormula = attributeValue == "1";
                            break;
                        case "l":
                            IsNewLevelDependency = attributeValue == "1";
                            break;
                        case "r":
                            CellReference = attributeValue;
                            break;
                        case "s":
                            IsChildChain = attributeValue == "1";
                            break;
                        case "t":
                            IsNewThread = attributeValue == "1";
                            break;
                    }
                }
            }
        }

        private class CellFormula
        {
            public bool CalculateCell { get; }
            public bool AlwaysCalculateArray { get; }
            public string RangeOfCells { get; }
            public string Formula { get; }
            public CellFormulaType FormulaType { get; }

            public CellFormula(Dictionary<string, string> attributes, string formula)
            {
                foreach (var attribute in attributes)
                {
                    var key = attribute.Key;
                    var value = attribute.Value;

                    if (key == "t")
                        FormulaType = ParseCellFormulaType(value);
                    if (key == "ref")
                        RangeOfCells = value;
                    if (key == "ca")
                        CalculateCell = value == "1";
                    if (key == "aca")
                        AlwaysCalculateArray = value == "1";
                }

                Formula = formula;
            }

            private static CellFormulaType ParseCellFormulaType(string s)
            {
                if (s == "array")
                    return CellFormulaType.Array;
                if (s == "dataTable")
                    return CellFormulaType.TableFormula;
                if (s == "normal")
                    return CellFormulaType.Normal;
                if (s == "shared")
                    return CellFormulaType.SharedFormula;

                // default value
                return CellFormulaType.Normal;
            }
        }

        #endregion Private Classes
    }
}