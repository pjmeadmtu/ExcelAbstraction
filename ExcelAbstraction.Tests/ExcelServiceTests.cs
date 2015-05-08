using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestHelper;

namespace ExcelAbstraction.Tests
{
	public abstract class ExcelServiceTests
	{
		protected readonly IExcelService ExcelService;

		protected Workbook DiskWorkbook, MemoryWorkbook;

		readonly string _fileName, _newFileName = Guid.NewGuid().ToString();

		readonly Validation[] _validations =
		{
			new Validation
			{
				Range = new Range
				{
					RowStart = 0,
					ColumnStart = 0,
					ColumnEnd = 0
				},
				List = new[] { "1871.02", "1871.03" }
			},
			new Validation
			{
				Range = new Range
				{
					RowStart = 0,
					RowEnd = 0,
					ColumnStart = 1,
					ColumnEnd = 1
				},
				List = new[] { "egergeg", "jhluliluil" }
			}
		};

		protected ExcelServiceTests(IExcelService excelService, string fileName)
		{
			ExcelService = excelService;
			_fileName = fileName;
		}

		public virtual void TestInitialize()
		{
			DiskWorkbook = ExcelService.ReadWorkbook(_fileName);
			MemoryWorkbook = new Workbook(new[]
			{
				new Worksheet(Guid.NewGuid().ToString(), 0, 2, new[]
				{
					new Row(0, new[]
					{
						new Cell(0, 0, Guid.NewGuid().ToString()),
						new Cell(0, 1, Guid.NewGuid().ToString())
					})
				})
			});
		}

		public virtual void TestCleanup()
		{
			File.Delete(_newFileName);
		}

		public virtual void ExcelService_OpenWorkbook_FileNotFound_ReturnsNull()
		{
			Assert.IsNull(ExcelService.ReadWorkbook(Guid.NewGuid().ToString()));
		}

		public virtual void ExcelService_CheckGrid()
		{
			foreach (var worksheet in DiskWorkbook.Worksheets)
			{
				var rows = worksheet.Rows.ToArray();

				for (int i = 0; i < rows.Length; i++)
				{
					var row = rows[i];

					Assert.AreEqual(i, row.Index);

					var cells = rows[i].Cells.ToArray();

					Assert.AreEqual(worksheet.ColumnCount, cells.Length);

					for (int j = 0; j < cells.Length; j++)
					{
						var cell = cells[j];
						if (cell == null) continue;

						Assert.AreEqual(i, cell.RowIndex);
						Assert.AreEqual(j, cell.ColumnIndex);
					}
				}
			}
		}

		public virtual void ExcelService_UsesCulture()
		{
			ExcelService.Format = new CultureInfo("en-US");
			var before = ExcelService.ReadWorkbook(_fileName);
			ExcelService.Format = new CultureInfo("de-DE");
			var after = ExcelService.ReadWorkbook(_fileName);
			Assert.IsFalse(ExcelCompare.Compare(before, after));
		}

		public virtual void ExcelService_IgnoresThreadCulture()
		{
			ExcelService.Format = new CultureInfo("en-US");
			var before = ExcelService.ReadWorkbook(_fileName);
			Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
			var after = ExcelService.ReadWorkbook(_fileName);
			Assert.IsTrue(ExcelCompare.Compare(before, after));
		}

		public virtual void ExcelService_WriteWorkbook_Xls()
		{
			ExcelService.WriteWorkbook(DiskWorkbook, ExcelVersion.Xls, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			Assert.IsTrue(ExcelCompare.Compare(DiskWorkbook, newWorkbook));
		}

		public virtual void ExcelService_WriteWorkbook_Xlsx()
		{
			ExcelService.WriteWorkbook(DiskWorkbook, ExcelVersion.Xlsx, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			Assert.IsTrue(ExcelCompare.Compare(DiskWorkbook, newWorkbook));
		}

		public virtual void ExcelService_AddValidations()
		{
			var oldValidations = MemoryWorkbook.Worksheets.First().Validations;
			foreach (var validation in _validations)
				oldValidations.Add(validation);

			ExcelService.WriteWorkbook(MemoryWorkbook, ExcelVersion.Xlsx, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			var newValidations = newWorkbook.Worksheets.First().Validations;

			AssertHelper.AreDeeplyEqual(oldValidations, newValidations);
		}

		public virtual void ExcelService_AddValidations_Hack()
		{
			ExcelService.WriteWorkbook(MemoryWorkbook, ExcelVersion.Xlsx, _newFileName);
			object oldWorkbook = ExcelService.GetWorkbook(_newFileName);
			File.Delete(_newFileName);

			ExcelService.AddValidations(oldWorkbook, 0, _validations);
			ExcelService.SaveWorkbook(oldWorkbook, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			var newValidations = newWorkbook.Worksheets.First().Validations;

			AssertHelper.AreDeeplyEqual(_validations, newValidations);
		}
	}
}