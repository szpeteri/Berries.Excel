using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace Berries.Excel.Test
{
    public class ExcelImportTests
    {
        private const string FileName = "simple.xlsx";
        private const string FirstSheetName = "FirstSheet";

        class Simple
        {
            public string ColumnA { get; set; }
            public string ColumnB { get; set; }

            public string ColumnC { get; set; }
        }


        [SetUp]
        public void Setup()
        {
        }

        [TearDown]
        public void TearDown()
        {
        }


        [Test]
        public void Test()
        {
            // Arrange
            var importer = new ExcelImport<Simple>
            {
                ColumnMapList = new List<ExcelImport<Simple>.ExcelColumnMap>
                {
                    new ExcelImport<Simple>.ExcelColumnMap(1, (x, value) => x.ColumnA = value),
                    new ExcelImport<Simple>.ExcelColumnMap(2, (x, value) => x.ColumnB = value),
                    new ExcelImport<Simple>.ExcelColumnMap(3, (x, value) => x.ColumnC = value)
                }
            };

            // Act
            var result = importer.ImportFromFile(FileName, FirstSheetName, 2).ToArray();

            // Assert
            Assert.AreEqual(2, result.Length);

            Assert.AreEqual("Data A", result[0].ColumnA);
            Assert.AreEqual("Data B", result[0].ColumnB);
            Assert.AreEqual("Data C", result[0].ColumnC);
            Assert.AreEqual("Third row", result[1].ColumnA);
            Assert.IsNull(result[1].ColumnB);
            Assert.AreEqual("Last cell", result[1].ColumnC);
        }
    }
}
