using System.Diagnostics;
using System.Text.RegularExpressions;

namespace SpreadsheetCalculator
{
    [DebuggerDisplay("{Column}{Row}")]
    public class CellReference
    {
        public int Row { get; }
        public string Column { get; }

        private readonly Regex _regex =
            new Regex($@"(?<{nameof(Column)}>[A-Z]+)(?<{nameof(Row)}>[0-9]+)", RegexOptions.Compiled);

        public CellReference(string cellReference)
        {
            var match = _regex.Match(cellReference);
            if (match.Success)
            {
                Row = int.Parse(match.Groups[nameof(Row)].Value);
                Column = match.Groups[nameof(Column)].Value;
            }
        }
    }
}