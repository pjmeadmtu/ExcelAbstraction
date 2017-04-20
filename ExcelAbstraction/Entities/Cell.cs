using System.Diagnostics;

namespace ExcelAbstraction.Entities
{
	[DebuggerDisplay("{Value}")]
	public class Cell
	{
		public int RowIndex { get; private set; }
		public int ColumnIndex { get; private set; }
		public string Value { get; private set; }
		public string DataFormat { get; private set; }

		public Cell(int rowIndex, int columnIndex, string value, string dataFormat = "")
		{
			RowIndex = rowIndex;
			ColumnIndex = columnIndex;
			Value = value;
			DataFormat = dataFormat;
		}
	}
}