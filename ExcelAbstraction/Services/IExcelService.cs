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
		object GetWorkbook(Stream stream);
		void SaveWorkbook(object workbook, string path);
		void SaveWorkbook(object workbook, Stream stream);
		void AddNames(object workbook, ExcelVersion version, params NamedRange[] names);
		void AddRows(object workbook, string sheetName, params Row[] rows);
		void AddValidations(object workbook, string sheetName, ExcelVersion version, params DataValidation[] validations);
	}
}