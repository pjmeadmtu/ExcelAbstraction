using System;
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
		public void ExcelHelper_ConvertRowNumberToIndex()
		{
			Assert.AreEqual(null, ExcelHelper.ConvertRowNumberToIndex("", ExcelVersion.Xls));
			Assert.AreEqual(0, ExcelHelper.ConvertRowNumberToIndex("1", ExcelVersion.Xls));
			Assert.AreEqual(0, ExcelHelper.ConvertRowNumberToIndex("1", ExcelVersion.Xlsx));
			Assert.AreEqual(65535, ExcelHelper.ConvertRowNumberToIndex("65536", ExcelVersion.Xls));
			Assert.AreEqual(1048575, ExcelHelper.ConvertRowNumberToIndex("1048576", ExcelVersion.Xlsx));
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertRowNumberToIndex_BelowMin_Xls()
		{
			ExcelHelper.ConvertRowNumberToIndex("0", ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertRowNumberToIndex_BelowMin_Xlsx()
		{
			ExcelHelper.ConvertRowNumberToIndex("0", ExcelVersion.Xlsx);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertRowNumberToIndex_AboveMin_Xls()
		{
			ExcelHelper.ConvertRowNumberToIndex("65537", ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertRowNumberToIndex_AboveMin_Xlsx()
		{
			ExcelHelper.ConvertRowNumberToIndex("1048577", ExcelVersion.Xlsx);
		}

		[TestMethod]
		public void ExcelHelper_ConvertColumnLettersToIndex()
		{
			Assert.AreEqual(null, ExcelHelper.ConvertColumnLettersToIndex("", ExcelVersion.Xls));
			Assert.AreEqual(0, ExcelHelper.ConvertColumnLettersToIndex("A", ExcelVersion.Xls));
			Assert.AreEqual(0, ExcelHelper.ConvertColumnLettersToIndex("A", ExcelVersion.Xlsx));
			Assert.AreEqual(255, ExcelHelper.ConvertColumnLettersToIndex("IV", ExcelVersion.Xls));
			Assert.AreEqual(16383, ExcelHelper.ConvertColumnLettersToIndex("XFD", ExcelVersion.Xlsx));
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertColumnLettersToIndex_AboveMin_Xls()
		{
			ExcelHelper.ConvertColumnLettersToIndex("IW", ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertColumnLettersToIndex_AboveMin_Xlsx()
		{
			ExcelHelper.ConvertColumnLettersToIndex("XFE", ExcelVersion.Xlsx);
		}

		[TestMethod]
		public void ExcelHelper_ConvertIndexToRowNumber()
		{
			Assert.AreEqual("1", ExcelHelper.ConvertIndexToRowNumber(0, ExcelVersion.Xls));
			Assert.AreEqual("1", ExcelHelper.ConvertIndexToRowNumber(0, ExcelVersion.Xlsx));
			Assert.AreEqual("65536", ExcelHelper.ConvertIndexToRowNumber(65535, ExcelVersion.Xls));
			Assert.AreEqual("1048576", ExcelHelper.ConvertIndexToRowNumber(1048575, ExcelVersion.Xlsx));
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToRowNumber_BelowMin_Xls()
		{
			ExcelHelper.ConvertIndexToRowNumber(-1, ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToRowNumber_BelowMin_Xlsx()
		{
			ExcelHelper.ConvertIndexToRowNumber(-1, ExcelVersion.Xlsx);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToRowNumber_AboveMax_Xls()
		{
			ExcelHelper.ConvertIndexToRowNumber(65536, ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToRowNumber_AboveMax_Xlsx()
		{
			ExcelHelper.ConvertIndexToRowNumber(1048576, ExcelVersion.Xlsx);
		}

		[TestMethod]
		public void ExcelHelper_ConvertIndexToColumnLetters()
		{
			Assert.AreEqual("A", ExcelHelper.ConvertIndexToColumnLetters(0, ExcelVersion.Xls));
			Assert.AreEqual("A", ExcelHelper.ConvertIndexToColumnLetters(0, ExcelVersion.Xlsx));
			Assert.AreEqual("IV", ExcelHelper.ConvertIndexToColumnLetters(255, ExcelVersion.Xls));
			Assert.AreEqual("XFD", ExcelHelper.ConvertIndexToColumnLetters(16383, ExcelVersion.Xlsx));
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToColumnLetters_BelowMin_Xls()
		{
			ExcelHelper.ConvertIndexToColumnLetters(-1, ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToColumnLetters_BelowMin_Xlsx()
		{
			ExcelHelper.ConvertIndexToColumnLetters(-1, ExcelVersion.Xlsx);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToColumnLetters_AboveMax_Xls()
		{
			ExcelHelper.ConvertIndexToColumnLetters(256, ExcelVersion.Xls);
		}

		[TestMethod, ExpectedException(typeof(InvalidOperationException))]
		public void ExcelHelper_ConvertIndexToColumnLetters_AboveMax_Xlsx()
		{
			ExcelHelper.ConvertIndexToColumnLetters(16384, ExcelVersion.Xlsx);
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
		public void ExcelHelper_ParseRange_SingleCell()
		{
			var expected = new Range
			{
				RowStart = 33,
				RowEnd = 33,
				ColumnStart = 2,
				ColumnEnd = 2
			};

			var actual = ExcelHelper.ParseRange("C34", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_IncludingSheetName()
		{
			var expected = new Range(33, 37, 2, 7, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$C$34:$H$38", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_NoRow()
		{
			var expected = new Range(null, null, 24, 24);

			var actual = ExcelHelper.ParseRange("Y", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_NoColumn()
		{
			var expected = new Range(6, 6, null, null);

			var actual = ExcelHelper.ParseRange("7", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_NoRows()
		{
			var expected = new Range(null, null, 2, 7, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$C:$H", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_NoColumns()
		{
			var expected = new Range(64, 343, null, null, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$65:$344", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_DetectAllRows_Xls()
		{
			var expected = new Range(null, null, 2, 7, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$C1:$H65536", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_DetectAllRows_Xlsx()
		{
			var expected = new Range(null, null, 2, 7, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$C1:$H1048576", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_DetectAllColumns_Xls()
		{
			var expected = new Range(64, 343, null, null, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$A$65:$IV$344", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_DetectAllColumns_Xlsx()
		{
			var expected = new Range(64, 343, null, null, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$A$65:$XFD$344", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_Max_Xls()
		{
			var expected = new Range(2, null, 5, null, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$F$3:$IV$65536", ExcelVersion.Xls);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_ParseRange_Max_Xlsx()
		{
			var expected = new Range(2, null, 5, null, "testSheet34");

			var actual = ExcelHelper.ParseRange("testSheet34!$F$3:$XFD$1048576", ExcelVersion.Xlsx);

			AssertHelper.AreDeeplyEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_RangeToString()
		{
			var expected = "testSheet34!$I$23:$X$45";

			var actual = ExcelHelper.RangeToString(new Range(22, 44, 8, 23, "testSheet34"), ExcelVersion.Xlsx);

			Assert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_RangeToString_OneCell()
		{
			var expected = "testSheet34!$I$23";

			var actual = ExcelHelper.RangeToString(new Range(22, 22, 8, 8, "testSheet34"), ExcelVersion.Xlsx);

			Assert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_RangeToString_TwoNulls()
		{
			var expected = "testSheet34!$65:$344";

			var actual = ExcelHelper.RangeToString(new Range(64, 343, null, null, "testSheet34"), ExcelVersion.Xls);

			Assert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_RangeToString_OneNull()
		{
			var expected = "testSheet34!$I:$XFD";

			var actual = ExcelHelper.RangeToString(new Range(null, null, 8, null, "testSheet34"), ExcelVersion.Xlsx);

			Assert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ExcelHelper_GetRowMax()
		{
			Assert.AreEqual(65536, ExcelHelper.GetRowMax(ExcelVersion.Xls));
			Assert.AreEqual(1048576, ExcelHelper.GetRowMax(ExcelVersion.Xlsx));
		}

		[TestMethod]
		public void ExcelHelper_GetColumnMax()
		{
			Assert.AreEqual(256, ExcelHelper.GetColumnMax(ExcelVersion.Xls));
			Assert.AreEqual(16384, ExcelHelper.GetColumnMax(ExcelVersion.Xlsx));
		}

		[TestMethod, ExpectedException(typeof(InvalidEnumArgumentException))]
		public void ExcelHelper_GetRowMax_InvalidEnum()
		{
			ExcelHelper.GetRowMax((ExcelVersion)(-1));
		}
	}
}