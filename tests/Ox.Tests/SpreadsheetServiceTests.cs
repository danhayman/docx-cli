using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ox.Tests;

public class SpreadsheetServiceTests
{
    [Fact]
    public void GetParagraphs_SingleSheet_ExtractsRows()
    {
        var path = CreateTestXlsx(new[] { new[] { "Alice", "30" }, new[] { "Bob", "25" } });
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var paragraphs = Ox.Core.SpreadsheetService.GetParagraphs(doc);

            Assert.Equal(2, paragraphs.Count);
            Assert.Equal("Alice\t30", paragraphs[0].Text);
            Assert.Equal("Bob\t25", paragraphs[1].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_SharedStrings_ResolvesCorrectly()
    {
        var path = CreateTestXlsxWithSharedStrings(
            new[] { "Hello", "World" },
            new[] { new[] { 0, 1 }, new[] { 1, 0 } });
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var paragraphs = Ox.Core.SpreadsheetService.GetParagraphs(doc);

            Assert.Equal(2, paragraphs.Count);
            Assert.Equal("Hello\tWorld", paragraphs[0].Text);
            Assert.Equal("World\tHello", paragraphs[1].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_SheetFilter_OnlyReturnsMatchingSheet()
    {
        var path = CreateMultiSheetXlsx(
            ("Sales", new[] { new[] { "Q1", "100" } }),
            ("Costs", new[] { new[] { "Q1", "50" } }));
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var paragraphs = Ox.Core.SpreadsheetService.GetParagraphs(doc, "Sales");

            Assert.Single(paragraphs);
            Assert.Equal("Q1\t100", paragraphs[0].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_AllSheets_ReturnsAll()
    {
        var path = CreateMultiSheetXlsx(
            ("Sales", new[] { new[] { "Q1", "100" } }),
            ("Costs", new[] { new[] { "Q1", "50" } }));
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var paragraphs = Ox.Core.SpreadsheetService.GetParagraphs(doc);

            Assert.Equal(2, paragraphs.Count);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetSheetNames_ReturnsAllNames()
    {
        var path = CreateMultiSheetXlsx(
            ("Alpha", new[] { new[] { "a" } }),
            ("Beta", new[] { new[] { "b" } }));
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var names = Ox.Core.SpreadsheetService.GetSheetNames(doc);

            Assert.Equal(2, names.Count);
            Assert.Equal("Alpha", names[0]);
            Assert.Equal("Beta", names[1]);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_EmptySheet_ReturnsEmpty()
    {
        var path = CreateTestXlsx(Array.Empty<string[]>());
        try
        {
            using var doc = SpreadsheetDocument.Open(path, false);
            var paragraphs = Ox.Core.SpreadsheetService.GetParagraphs(doc);

            Assert.Empty(paragraphs);
        }
        finally { File.Delete(path); }
    }

    // -- Helpers --

    private static string CreateTestXlsx(string[][] rows)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx_test_{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        foreach (var row in rows)
        {
            var sheetRow = new Row();
            foreach (var value in row)
            {
                sheetRow.Append(new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(value))
                });
            }
            sheetData.Append(sheetRow);
        }

        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = new Sheets();
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
        workbookPart.Workbook.Append(sheets);
        workbookPart.Workbook.Save();

        return path;
    }

    private static string CreateTestXlsxWithSharedStrings(string[] strings, int[][] rowIndices)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx_test_{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Create shared string table
        var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();
        foreach (var s in strings)
            sst.Append(new SharedStringItem(new Text(s)));
        sstPart.SharedStringTable = sst;

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        foreach (var rowIdx in rowIndices)
        {
            var sheetRow = new Row();
            foreach (var idx in rowIdx)
            {
                sheetRow.Append(new Cell
                {
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(idx.ToString())
                });
            }
            sheetData.Append(sheetRow);
        }

        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = new Sheets();
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
        workbookPart.Workbook.Append(sheets);
        workbookPart.Workbook.Save();

        return path;
    }

    private static string CreateMultiSheetXlsx(params (string Name, string[][] Rows)[] sheetDefs)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx_test_{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sheets = new Sheets();
        uint sheetId = 1;

        foreach (var (name, rows) in sheetDefs)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();

            foreach (var row in rows)
            {
                var sheetRow = new Row();
                foreach (var value in row)
                {
                    sheetRow.Append(new Cell
                    {
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(new Text(value))
                    });
                }
                sheetData.Append(sheetRow);
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);

            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = name
            });
        }

        workbookPart.Workbook.Append(sheets);
        workbookPart.Workbook.Save();

        return path;
    }
}
