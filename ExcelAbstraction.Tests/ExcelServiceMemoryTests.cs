using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using ExcelAbstraction.Entities;
using ExcelAbstraction.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestHelper;

namespace ExcelAbstraction.Tests
{
	public abstract class ExcelServiceMemoryTests
	{
		protected readonly IExcelService ExcelService;

		protected Workbook MemoryWorkbook, MemoryToDiskWorkbook;

		protected object MemoryToDiskObject;

		readonly ExcelVersion _version;
		readonly string _newFileName = Guid.NewGuid().ToString();

		readonly DataValidation[] _validations =
		{
			new DataValidation
			{
				Range = new Range
				{
					RowStart = 1,
					ColumnStart = 0,
					ColumnEnd = 0
				},
				List = new[] { "1871.02", "1871.03" }
			},
			new DataValidation
			{
				Range = new Range
				{
					RowStart = 1,
					RowEnd = 2,
				},
				List = new[] { "1871.02", "1871.03" }
			},
			new DataValidation
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

		protected ExcelServiceMemoryTests(IExcelService excelService, ExcelVersion version)
		{
			ExcelService = excelService;
			_version = version;
		}

		public virtual void TestInitialize()
		{
			var collection = new Collection<string>();
			for (int i = 0; i < 100; i++)
				collection.Add(Guid.NewGuid().ToString());
			MemoryWorkbook = new Workbook(new[]
			{
				new Worksheet("testSheet5", 0, 2, new[]
				{
					new Row(0, new[]
					{
						new Cell(0, 0, Guid.NewGuid().ToString()),
						new Cell(0, 1, Guid.NewGuid().ToString())
					})
				}),
				new Worksheet("testName4", 1, 1, collection.Select((c, i) => new Row(i, new[] { new Cell(i, 0, c) })))
				{
					IsHidden = true
				}
			});

			ExcelService.WriteWorkbook(MemoryWorkbook, _version, _newFileName);
			MemoryToDiskObject = ExcelService.GetWorkbook(_newFileName);
			MemoryToDiskWorkbook = ExcelService.ReadWorkbook(_newFileName);
			File.Delete(_newFileName);
		}

		public virtual void TestCleanup()
		{
			File.Delete(_newFileName);
		}

		public virtual void ExcelService_AddValidations()
		{
			var oldValidations = MemoryWorkbook.Worksheets.First().Validations;
			foreach (var validation in _validations)
				oldValidations.Add(validation);

			ExcelService.WriteWorkbook(MemoryWorkbook, _version, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			var newValidations = newWorkbook.Worksheets.First().Validations;

			AssertHelper.AreDeeplyEqual(oldValidations, newValidations);
		}

		public virtual void ExcelService_AddValidations_Hack()
		{
			ExcelService.AddValidations(MemoryToDiskObject, "testSheet5", _version, _validations);
			ExcelService.SaveWorkbook(MemoryToDiskObject, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			var newValidations = newWorkbook.Worksheets.First().Validations;

			AssertHelper.AreDeeplyEqual(_validations, newValidations);
		}

		public virtual void ExcelService_AddValidations_NullListThrows()
		{
			var oldValidations = new[] { new DataValidation() };

			ExcelService.AddValidations(MemoryToDiskObject, "testSheet5", _version, oldValidations);
		}

		public virtual void ExcelService_AddValidations_EmptyListThrows()
		{
			var oldValidations = new[] { new DataValidation { List = new Collection<string>() } };

			ExcelService.AddValidations(MemoryToDiskObject, "testSheet5", _version, oldValidations);
		}

		public virtual void ExcelService_Worksheet_IsHidden()
		{
			Assert.AreEqual(true, MemoryToDiskWorkbook.Worksheets.Last().IsHidden);
		}

		public virtual void ExcelService_AddValidations_LotsOfItems()
		{
			var oldNames = new[]
			{
				new NamedRange
				{
					Name = "testName3",
					Range = new Range(0, 99, 0, 0, "testName4")
				}
			};

			var oldValidations = new[]
			{
				new DataValidation
				{
					Name = "testName3",
					Type = DataValidationType.Formula,
					Range = new Range()
				}
			};

			ExcelService.AddNames(MemoryToDiskObject, _version, oldNames);
			ExcelService.AddValidations(MemoryToDiskObject, "testSheet5", _version, oldValidations);
			ExcelService.SaveWorkbook(MemoryToDiskObject, _newFileName);
			var newWorkbook = ExcelService.ReadWorkbook(_newFileName);
			var newNames = newWorkbook.Names;
			var newValidations = newWorkbook.Worksheets.First().Validations;

			AssertHelper.AreDeeplyEqual(oldNames, newNames);
			AssertHelper.AreDeeplyEqual(oldValidations, newValidations);
		}
	}
}