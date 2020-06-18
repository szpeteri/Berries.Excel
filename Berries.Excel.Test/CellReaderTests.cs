using NUnit.Framework;

namespace Berries.Excel.Test
{
    public class CellReaderTests
    {
        private const string FileName = "simple.xlsx";
        private const string FirstSheetName = "FirstSheet";
        private const string SecondSheetName = "SecondSheet";

        private Package _package;

        [SetUp]
        public void Setup()
        {
            _package = new Package(FileName);
        }


        [TearDown]
        public void TearDown()
        {
            _package?.Dispose();
            _package = null;
        }

        [TestCase("a1", "Header A")]
        [TestCase("b1", "Header B")]
        [TestCase("c1", "Header C")]
        [TestCase("a2", "Data A")]
        [TestCase("b2", "Data B")]
        [TestCase("c2", "Data C")]
        public void ReadCells(string address, string value)
        {
            // Arrange
            var reader = CellReader.Create(_package.Workbook[FirstSheetName]);

            // Act
            var cell = reader.GetCell(address);

            // Assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(value, cell.Value);

            reader.Dispose();
        }

        [TestCase(1, 1, "Header A")]
        [TestCase(1, 2, "Header B")]
        [TestCase(1, 3, "Header C")]
        [TestCase(2, 1, "Data A")]
        [TestCase(2, 2, "Data B")]
        [TestCase(2, 3, "Data C")]
        public void ReadCellsByCoordinates(int row, int col, string value)
        {
            // Arrange
            var reader = CellReader.Create(_package.Workbook[FirstSheetName]);

            // Act
            var cell = reader.GetCell(row, col);

            // Assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(value, cell.Value);

            reader.Dispose();
        }

        [TestCase(1, 1, "A1")]
        [TestCase(2, 1, "A2")]
        [TestCase(1, 2, "B1")]
        [TestCase(2, 2, "B2")]
        public void ColumnName(int row, int col, string address)
        {
            // Arrange
            var reader = CellReader.Create(_package.Workbook[FirstSheetName]);

            // Act
            var cell = reader.GetCell(row, col);

            // Assert
            Assert.AreEqual(col, cell.ColumnIndex);
        }
    }
}