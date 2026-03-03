using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Ox.Tests;

public static class TestHelper
{
    public static string CreateTestDocx(params string[] paragraphTexts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docx_test_{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();

        foreach (var text in paragraphTexts)
        {
            var para = new Paragraph();
            if (!string.IsNullOrEmpty(text))
            {
                para.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            }
            body.Append(para);
        }

        mainPart.Document.Append(body);
        mainPart.Document.Save();

        return path;
    }

    public static string CreateMultiRunDocx(string[] runTexts, RunProperties? props = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docx_test_{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        var para = new Paragraph();

        foreach (var text in runTexts)
        {
            var run = new Run();
            if (props != null)
                run.Append(props.CloneNode(true));
            run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            para.Append(run);
        }

        body.Append(para);
        mainPart.Document.Append(body);
        mainPart.Document.Save();

        return path;
    }

    public static string CreateDocxWithTable(string[] paraTexts, string[][] tableRows)
    {
        var path = Path.Combine(Path.GetTempPath(), $"docx_test_{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();

        foreach (var text in paraTexts)
        {
            body.Append(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
        }

        var table = new Table();
        foreach (var row in tableRows)
        {
            var tr = new TableRow();
            foreach (var cell in row)
            {
                var tc = new TableCell(new Paragraph(new Run(new Text(cell) { Space = SpaceProcessingModeValues.Preserve })));
                tr.Append(tc);
            }
            table.Append(tr);
        }
        body.Append(table);

        mainPart.Document.Append(body);
        mainPart.Document.Save();

        return path;
    }

    public static string ReadDocxText(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        return string.Join("\n", body.Elements<Paragraph>().Select(p => p.InnerText));
    }

    public static List<string> ReadDocxParagraphs(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        return body.Elements<Paragraph>().Select(p => p.InnerText).ToList();
    }
}
