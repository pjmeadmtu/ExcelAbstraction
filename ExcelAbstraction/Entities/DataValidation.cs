using System.Collections.Generic;

namespace ExcelAbstraction.Entities
{
	public class DataValidation
	{
		public Range Range { get; set; }
		public DataValidationType Type { get; set; }
		public ICollection<string> List { get; set; }
		public string Name { get; set; }
	}
}