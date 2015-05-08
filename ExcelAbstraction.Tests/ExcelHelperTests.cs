using System.ComponentModel;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestHelper;

namespace ExcelAbstraction.Tests
{
	[TestClass]
	public class ExcelHelperTests
	{
		[TestMethod]
		public void ExcelHelper_ConvertColumnLettersToIndex()
		{
			Assert.AreEqual(0, ExcelHelper.ConvertColumnLettersToIndex("A"));
			Assert.AreEqual(730, ExcelHelper.ConvertColumnLettersToIndex("ABC"));
			Assert.AreEqual(17014, ExcelHelper.ConvertColumnLettersToIndex("YDK"));
			Assert.AreEqual(3940218, ExcelHelper.ConvertColumnLettersToIndex("HPDRW"));
		}

		[TestMethod]
		public void ExcelHelper_ParseRange()
		{
			var expected = new Range
			{
				RowStart = 33,
				RowEnd = 37,
				ColumnStart = 0,
				ColumnEnd = 7
			};

			var actual = ExcelHelper.ParseRange("A34:H38", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_RowMax()
		{
			Assert.AreEqual(65534, ExcelHelper.ParseRange("A1:A65535", ExcelVersion.Xls).RowEnd);
			Assert.IsNull(ExcelHelper.ParseRange("A1:A65536", ExcelVersion.Xls).RowEnd);
			Assert.IsNull(ExcelHelper.ParseRange("A1:A65537", ExcelVersion.Xls).RowEnd);

			Assert.AreEqual(1048574, ExcelHelper.ParseRange("A1:A1048575", ExcelVersion.Xlsx).RowEnd);
			Assert.IsNull(ExcelHelper.ParseRange("A1:A1048576", ExcelVersion.Xlsx).RowEnd);
			Assert.IsNull(ExcelHelper.ParseRange("A1:A1048577", ExcelVersion.Xlsx).RowEnd);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_SingleCell()
		{
			var expected = new Range
			{
				RowStart = 33,
				RowEnd = 33,
				ColumnStart = 0,
				ColumnEnd = 0
			};

			var actual = ExcelHelper.ParseRange("A34", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_GetRowMax()
		{
			Assert.AreEqual(ExcelHelper.GetRowMax(ExcelVersion.Xls), 65535);
			Assert.AreEqual(ExcelHelper.GetRowMax(ExcelVersion.Xlsx), 1048575);
		}

		[TestMethod, ExpectedException(typeof(InvalidEnumArgumentException))]
		public void ExcelHelper_GetRowMax_InvalidEnum()
		{
			ExcelHelper.GetRowMax((ExcelVersion)(-1));
		}
	}
}