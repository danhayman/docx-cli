using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using Ox.Core;
using Ox.Output;

namespace Ox.Commands;

public static class InfoCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string>("file") { Description = "Path to .docx, .pptx, or .xlsx file" };
        var command = new Command("info") { Description = "Show document metadata (title, author, word count, etc.)" };
        command.Arguments.Add(fileArg);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var file = parseResult.GetValue(fileArg)!;
            var ext = Path.GetExtension(file).ToLowerInvariant();

            return ext switch
            {
                ".docx" => ShowDocxInfo(file),
                ".pptx" => ShowPptxInfo(file),
                ".xlsx" => ShowXlsxInfo(file),
                _ => ShowUnsupported(ext)
            };
        });

        return command;
    }

    private static Task<int> ShowDocxInfo(string file)
    {
        using var doc = DocumentService.OpenRead(file);
        var body = doc.MainDocumentPart!.Document.Body!;
        var props = doc.PackageProperties;
        var fileInfo = new FileInfo(file);

        var wordCount = DocumentService.CountWords(body);
        var paragraphs = DocumentService.GetParagraphs(body);
        var estimatedPages = Math.Max(1, (int)Math.Ceiling(wordCount / 250.0));

        TextFormatter.WriteTable("", new (string, string)[]
        {
            ("File:", fileInfo.Name),
            ("Format:", "Word (.docx)"),
            ("Size:", FormatSize(fileInfo.Length)),
            ("Pages:", $"{estimatedPages} (estimated)"),
            ("Paragraphs:", paragraphs.Count.ToString()),
            ("Words:", wordCount.ToString("N0")),
            ("Author:", props.Creator ?? "(unknown)"),
            ("Created:", props.Created?.ToString("o") ?? "(unknown)"),
            ("Modified:", props.Modified?.ToString("o") ?? "(unknown)")
        });

        return Task.FromResult(0);
    }

    private static Task<int> ShowPptxInfo(string file)
    {
        using var doc = PresentationDocument.Open(file, false);
        var props = doc.PackageProperties;
        var fileInfo = new FileInfo(file);

        var slideCount = PresentationService.CountSlides(doc);
        var wordCount = PresentationService.CountWords(doc);

        TextFormatter.WriteTable("", new (string, string)[]
        {
            ("File:", fileInfo.Name),
            ("Format:", "PowerPoint (.pptx)"),
            ("Size:", FormatSize(fileInfo.Length)),
            ("Slides:", slideCount.ToString()),
            ("Words:", wordCount.ToString("N0")),
            ("Author:", props.Creator ?? "(unknown)"),
            ("Created:", props.Created?.ToString("o") ?? "(unknown)"),
            ("Modified:", props.Modified?.ToString("o") ?? "(unknown)")
        });

        return Task.FromResult(0);
    }

    private static Task<int> ShowXlsxInfo(string file)
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        var props = doc.PackageProperties;
        var fileInfo = new FileInfo(file);

        var sheetNames = SpreadsheetService.GetSheetNames(doc);
        var wordCount = SpreadsheetService.CountWords(doc);

        var rows = new List<(string, string)>
        {
            ("File:", fileInfo.Name),
            ("Format:", "Excel (.xlsx)"),
            ("Size:", FormatSize(fileInfo.Length)),
            ("Sheets:", sheetNames.Count.ToString()),
            ("Words:", wordCount.ToString("N0")),
            ("Author:", props.Creator ?? "(unknown)"),
            ("Created:", props.Created?.ToString("o") ?? "(unknown)"),
            ("Modified:", props.Modified?.ToString("o") ?? "(unknown)")
        };

        for (int i = 0; i < sheetNames.Count; i++)
            rows.Add(($"  Sheet {i + 1}:", sheetNames[i]));

        TextFormatter.WriteTable("", rows);

        return Task.FromResult(0);
    }

    private static Task<int> ShowUnsupported(string ext)
    {
        Console.Error.WriteLine($"error: unsupported file format: {ext}");
        return Task.FromResult(1);
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        return $"{bytes / (1024.0 * 1024.0):F1} MB";
    }
}
