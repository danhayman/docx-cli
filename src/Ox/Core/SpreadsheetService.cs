using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ox.Core;

public static class SpreadsheetService
{
    public static List<ParagraphInfo> GetParagraphs(SpreadsheetDocument doc, string? sheetName = null)
    {
        var result = new List<ParagraphInfo>();
        int index = 1;

        var workbookPart = doc.WorkbookPart
            ?? throw new InvalidOperationException("invalid .xlsx: no workbook part");

        var sharedStrings = BuildSharedStrings(workbookPart);

        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>()
            ?? Enumerable.Empty<Sheet>();

        foreach (var sheet in sheets)
        {
            if (sheetName != null
                && !string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                continue;

            var relId = sheet.Id?.Value
                ?? throw new InvalidOperationException($"sheet '{sheet.Name}' missing relationship ID");

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relId);

            foreach (var row in worksheetPart.Worksheet.Descendants<Row>())
            {
                var cells = row.Elements<Cell>()
                    .Select(c => ResolveCellValue(c, sharedStrings));
                var tsv = string.Join("\t", cells);
                result.Add(new ParagraphInfo(index++, null, tsv));
            }
        }

        return result;
    }

    public static List<string> GetSheetNames(SpreadsheetDocument doc)
    {
        var workbookPart = doc.WorkbookPart
            ?? throw new InvalidOperationException("invalid .xlsx: no workbook part");

        return (workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .Select(s => s.Name?.Value ?? "(unnamed)")
            ?? Enumerable.Empty<string>())
            .ToList();
    }

    public static int CountWords(SpreadsheetDocument doc)
    {
        int count = 0;
        foreach (var para in GetParagraphs(doc))
        {
            if (!string.IsNullOrWhiteSpace(para.Text))
                count += para.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        return count;
    }

    private static string[] BuildSharedStrings(WorkbookPart workbookPart)
    {
        var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sst == null) return [];

        return sst.Elements<SharedStringItem>()
            .Select(item => item.InnerText)
            .ToArray();
    }

    private static string ResolveCellValue(Cell cell, string[] sharedStrings)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText ?? "";

        var raw = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(raw, out int idx)
            && idx >= 0 && idx < sharedStrings.Length)
        {
            return sharedStrings[idx];
        }

        return raw;
    }
}
