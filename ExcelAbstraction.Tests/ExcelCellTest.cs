using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelAbstraction.Entities;

namespace ExcelAbstraction.Tests
{
	[TestClass]
	public class ExcelCellTest
	{
		[TestMethod]
		public void TestCellDataFormat()
		{
			const string dateFormat = "MM/dd/yyyy";
			var cell = new Cell(0, 2, "3/12/2017", dateFormat);

			Assert.AreEqual(dateFormat, cell.DataFormat);
		}

		[TestMethod]
		public void TestCellDataFormatEmpty()
		{
			var cell = new Cell(0, 2, "$40,000,000");
			Assert.AreEqual(cell.DataFormat, string.Empty);
		}
	}
}