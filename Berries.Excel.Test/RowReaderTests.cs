using NUnit.Framework;

namespace Berries.Excel.Test
{
    public class RowReaderTests
    {
        private const string FileName = "simple.xlsx";
        private const string FirstSheetName = "FirstSheet";

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

        [TestCase(0, "A1", "Header A")]
        [TestCase(1, "B1", "Header B")]
        [TestCase(2, "C1", "Header C")]
        public void FirstRow(int cellIndex, string address, string value)
        {
            // Arrange
            var reader = RowReader.Create(_package.Workbook[FirstSheetName]);

            // Act
            var result = reader.Read();

            // Assert
            Assert.IsTrue(result);
            Assert.IsNotNull(reader.Row);
            Assert.AreEqual(3, reader.Row.Cells.Length);

            Assert.AreEqual(address, reader.Row.Cells[cellIndex].Address);
            Assert.AreEqual(value, reader.Row.Cells[cellIndex].Value);
        }

        [TestCase(0, "A2", "Data A")]
        [TestCase(1, "B2", "Data B")]
        [TestCase(2, "C2", "Data C")]
        public void SecondRow(int cellIndex, string address, string value)
        {
            // Arrange
            var reader = RowReader.Create(_package.Workbook[FirstSheetName]);

            // Act
            reader.Read();
            var result = reader.Read();

            // Assert
            Assert.IsTrue(result);
            Assert.IsNotNull(reader.Row);
            Assert.AreEqual(3, reader.Row.Cells.Length);

            Assert.AreEqual(address, reader.Row.Cells[cellIndex].Address);
            Assert.AreEqual(value, reader.Row.Cells[cellIndex].Value);
        }
    }
}