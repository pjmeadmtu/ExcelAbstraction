using System;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelAbstraction.Entities;

namespace ExcelAbstraction.Helpers
{
	public static class ExcelHelper
	{
		public const int
			RowMaxXls = 65536,
			RowMaxXlsx = 1048576,
			ColumnMaxXls = 256,
			ColumnMaxXlsx = 16384;

		static readonly Regex
			ColumnRegex = new Regex(@"([a-zA-Z]+)"),
			RowRegex = new Regex(@"(\d+)");

		public static int? ConvertRowNumberToIndex(string rowNumber, ExcelVersion version)
		{
			if (rowNumber == "") return null;

			int row = int.Parse(rowNumber);
			if (row < 1)
				throw new InvalidOperationException("row number must be greater than zero");

			//if (row > GetRowMax(version)) return null;
			int rowMax = GetRowMax(version);
			if (row > rowMax)
				throw new InvalidOperationException("row number must be less than or equal to " + rowMax);

			int index = row - 1;
			return index;
		}

		public static int? ConvertColumnLettersToIndex(string columnLetters, ExcelVersion version)
		{
			if (columnLetters == "") return null;

			int columnMax = GetColumnMax(version);

			int column = 0;
			for (int i = 0; i < columnLetters.Length; i++)
			{
				int num = columnLetters[columnLetters.Length - 1 - i] - 64;
				column += num * (int)Math.Pow(26, i);

				//if (column > columnMax) return null;
				if (column > columnMax)
					throw new InvalidOperationException("column letters must be less than or equal to " + ConvertIndexToColumnLetters(columnMax - 1, version));
			}

			int index = column - 1;
			return index;
		}

		public static string ConvertIndexToRowNumber(int index, ExcelVersion version)
		{
			if (index < 0)
				throw new InvalidOperationException("row index must be greater than or equal to zero");

			int rowNumber = index + 1;
			int rowMax = GetRowMax(version);
			if (rowNumber > rowMax)
				throw new InvalidOperationException("row index must be less than " + rowMax);

			return rowNumber.ToString();
		}

		public static string ConvertIndexToColumnLetters(int index, ExcelVersion version)
		{
			if (index < 0)
				throw new InvalidOperationException("column index must be greater than or equal to zero");

			int columnNumber = index + 1;
			int columnMax = GetColumnMax(version);
			if (columnNumber > columnMax)
				throw new InvalidOperationException("column index must be less than " + columnMax);

			var columnLetters = new StringBuilder();
			do
			{
				int rem = columnNumber % 26;
				columnLetters.Append((char)(rem + 64));
			} while ((columnNumber = columnNumber / 26) != 0);

			char[] charArray = columnLetters.ToString().ToCharArray();
			Array.Reverse(charArray);
			return new string(charArray);
		}

		public static Range ParseRange(string rangeString, ExcelVersion version)
		{
			var range = new Range();

			string[] split = rangeString.Split('!');

			if (split.Length > 1)
			{
				range.SheetName = split[0];
				split[1] = split[1].Replace("$", "");
			}

			var cells = split.Last().Split(':').Select(r => new
			{
				ColumnIndex = ConvertColumnLettersToIndex(ColumnRegex.Match(r).Value, version),
				RowIndex = ConvertRowNumberToIndex(RowRegex.Match(r).Value, version)
			}).ToArray();

			var start = cells[0];
			var end = cells.Length == 1 ? start : cells[1];

			Func<int?, int?, RowColumn, bool> isAll = (startIndex, endIndex, rowColumn) =>
				startIndex == 0 && endIndex == GetMax(rowColumn, version) - 1;

			Func<int?, RowColumn, int?> getMax = (index, rowColumn) =>
				index == GetMax(rowColumn, version) - 1 ? null : index;

			if (!isAll(start.RowIndex, end.RowIndex, RowColumn.Row))
			{
				range.RowStart = start.RowIndex;
				range.RowEnd = getMax(end.RowIndex, RowColumn.Row);
			}
			if (!isAll(start.ColumnIndex, end.ColumnIndex, RowColumn.Column))
			{
				range.ColumnStart = start.ColumnIndex;
				range.ColumnEnd = getMax(end.ColumnIndex, RowColumn.Column);
			}

			return range;
		}

		public static string RangeToString(Range range, ExcelVersion version)
		{
			string sheetName = range.SheetName != null ? range.SheetName + "!" : "";

			Func<object, string> prepend = s => (sheetName.Length > 0 ? "$" : "") + s;

			Func<int?, int?, Func<int, ExcelVersion, string>, RowColumn, Tuple<string, string>> getCellString = (start, end, convert, rowColumn) =>
			{
				string startString = "", endString = "";
				if (start != null || end != null)
				{
					startString = prepend(convert(start ?? 0, version));
					endString = prepend(convert(end ?? GetMax(rowColumn, version) - 1, version));
				}
				return new Tuple<string, string>(startString, endString);
			};

			Tuple<string, string> row = getCellString(range.RowStart, range.RowEnd, ConvertIndexToRowNumber, RowColumn.Row);
			Tuple<string, string> column = getCellString(range.ColumnStart, range.ColumnEnd, ConvertIndexToColumnLetters, RowColumn.Column);

			string first = column.Item1 + row.Item1;
			string second = column.Item2 + row.Item2;
			string rangeString = first == second ? first : string.Format("{0}:{1}", first, second);

			return sheetName + rangeString;
		}

		enum RowColumn { Row, Column }

		static int GetMax(RowColumn rowColumn, ExcelVersion version)
		{
			switch (rowColumn)
			{
				case RowColumn.Row:
					switch (version)
					{
						case ExcelVersion.Xls: return RowMaxXls;
						case ExcelVersion.Xlsx: return RowMaxXlsx;
					}
					break;
				case RowColumn.Column:
					switch (version)
					{
						case ExcelVersion.Xls: return ColumnMaxXls;
						case ExcelVersion.Xlsx: return ColumnMaxXlsx;
					}
					break;
			}
			throw new InvalidEnumArgumentException("version", (int)version, version.GetType());
		}

		public static int GetRowMax(ExcelVersion version)
		{
			return GetMax(RowColumn.Row, version);
		}

		public static int GetColumnMax(ExcelVersion version)
		{
			return GetMax(RowColumn.Column, version);
		}
	}
}