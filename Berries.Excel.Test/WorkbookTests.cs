using System.Linq;
using NUnit.Framework;

namespace Berries.Excel.Test
{
    public class WorkbookTests
    {
        private const string FileName = "simple.xlsx";
        private const string FirstSheetName = "FirstSheet";
        private const string SecondSheetName = "SecondSheet";

        private Package _package = null;

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

        [Test]
        public void CoreProperties()
        {
            // Arrange

            // Act

            // Assert
            Assert.AreEqual("Szilárd Péteri", _package.CoreProperties.CreatedBy);
            Assert.AreEqual("Szilárd Péteri", _package.CoreProperties.ModifedBy);
            Assert.AreEqual("2020-06-05T13:36:46Z", _package.CoreProperties.CreatedAt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }

        [Test]
        public void LoadWorkbookWithWorksheetNames()
        {
            // Arrange

            // Act

            // Assert
            Assert.AreEqual(2, _package.Workbook.Worksheets.Length);

            Assert.IsTrue(_package.Workbook.Worksheets.Any(x => x.Name == FirstSheetName));
            Assert.IsTrue(_package.Workbook.Worksheets.Any(x => x.Name == SecondSheetName));
        }

        [TestCase(FirstSheetName, "A1:C3")]
        [TestCase(SecondSheetName, "A1")]
        public void WorksheetDimensions(string sheetName, string dimension)
        {
            // Arrange

            // Act

            // Assert
            Assert.AreEqual(dimension, _package.Workbook[sheetName].Dimension);
        }
    }
}