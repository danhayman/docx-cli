using System.CommandLine;
using DocxCli.Core;
using DocxCli.Output;

namespace DocxCli.Commands;

public static class InfoCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string>("file") { Description = "Path to .docx file" };
        var command = new Command("info") { Description = "Show document metadata (title, author, word count)" };
        command.Arguments.Add(fileArg);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var file = parseResult.GetValue(fileArg)!;

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
                ("Size:", FormatSize(fileInfo.Length)),
                ("Pages:", $"{estimatedPages} (estimated)"),
                ("Paragraphs:", paragraphs.Count.ToString()),
                ("Words:", wordCount.ToString("N0")),
                ("Author:", props.Creator ?? "(unknown)"),
                ("Created:", props.Created?.ToString("o") ?? "(unknown)"),
                ("Modified:", props.Modified?.ToString("o") ?? "(unknown)")
            });

            return Task.FromResult(0);
        });

        return command;
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        return $"{bytes / (1024.0 * 1024.0):F1} MB";
    }
}
