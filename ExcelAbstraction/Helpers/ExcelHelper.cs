using System;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelAbstraction.Entities;

namespace ExcelAbstraction.Helpers
{
	public static class ExcelHelper
	{
		const int
			RowMaxXls = 65535,
			RowMaxXlsx = 1048575;

		static readonly Regex Regex = new Regex(@"([a-zA-Z]+)(\d+)");

		public static int ConvertColumnLettersToIndex(string columnLetters)
		{
			int columnNumber = 0;
			for (int i = 0; i < columnLetters.Length; i++)
			{
				int num = columnLetters[columnLetters.Length - 1 - i] - 64;
				columnNumber += num * (int)Math.Pow(26, i);
			}
			int index = columnNumber - 1;
			return index;
		}

		public static Range ParseRange(string rangeString, ExcelVersion version)
		{
			var cells = rangeString.Split(':').Select(r =>
			{
				GroupCollection groups = Regex.Match(r).Groups;
				return new
				{
					ColumnIndex = ConvertColumnLettersToIndex(groups[1].Value),
					RowIndex = int.Parse(groups[2].Value) - 1
				};
			}).ToArray();

			var start = cells[0];
			var end = cells.Length == 1 ? start : cells[1];

			var range = new Range
			{
				RowStart = start.RowIndex,
				ColumnStart = start.ColumnIndex,
				ColumnEnd = end.ColumnIndex
			};

			int rowEnd = end.RowIndex;
			if (rowEnd < GetRowMax(version))
				range.RowEnd = rowEnd;

			return range;
		}

		public static int GetRowMax(ExcelVersion version)
		{
			switch (version)
			{
				case ExcelVersion.Xls: return RowMaxXls;
				case ExcelVersion.Xlsx: return RowMaxXlsx;
			}
			throw new InvalidEnumArgumentException("version", (int)version, version.GetType());
		}
	}
}