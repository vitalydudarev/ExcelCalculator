using Xunit;

namespace SpreadsheetCalculator.Tests
{
    public class CellReferenceTests
    {
        [Fact]
        public void InstantiateCellReferenceWithCorrectInput()
        {
            var cellReference = new CellReference("ADFE4983");
            Assert.Equal("ADFE", cellReference.Column);
            Assert.Equal(4983, cellReference.Row);
        }
    }
}