using ExcelAbstraction.Entities;

namespace ExcelAbstraction.Services
{
	public interface IExcelService
	{
		Workbook ReadWorkbook(string path);
	}
}