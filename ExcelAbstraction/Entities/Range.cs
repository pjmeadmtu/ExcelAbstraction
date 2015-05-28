namespace ExcelAbstraction.Entities
{
	public class Range
	{
		public int? RowStart { get; set; }
		public int? RowEnd { get; set; }
		public int? ColumnStart { get; set; }
		public int? ColumnEnd { get; set; }
		public string SheetName { get; set; }

		public Range() { }

		public Range(int? rowStart, int? rowEnd, int? columnStart, int? columnEnd, string sheetName = null)
		{
			RowStart = rowStart;
			RowEnd = rowEnd;
			ColumnStart = columnStart;
			ColumnEnd = columnEnd;
			SheetName = sheetName;
		}
	}
}