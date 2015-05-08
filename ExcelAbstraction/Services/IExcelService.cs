using System;
using System.IO;
using ExcelAbstraction.Entities;

namespace ExcelAbstraction.Services
{
	public interface IExcelService
	{
		IFormatProvider Format { get; set; }
		Workbook ReadWorkbook(string path);
		Workbook ReadWorkbook(Stream stream);
		void WriteWorkbook(Workbook workbook, ExcelVersion version, string path);
		void WriteWorkbook(Workbook workbook, ExcelVersion version, Stream stream);
		object GetWorkbook(string path);
		void SaveWorkbook(object workbook, string path);
		void AddValidations(object workbook, int sheetIndex, params Validation[] validations);
	}
}