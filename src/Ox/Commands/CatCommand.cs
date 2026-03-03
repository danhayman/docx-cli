using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;
using Ox.Core;

namespace Ox.Commands;

public static class CatCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string[]>("file") { Description = "Path(s) or glob pattern(s) for .docx, .pptx, or .xlsx files", Arity = ArgumentArity.OneOrMore };
        var sheetOption = new Option<string?>("--sheet") { Description = "For .xlsx: only output rows from this sheet (default: all sheets)" };

        var command = new Command("cat") { Description = "Dump plain text from Office documents. Supports globs (e.g. '*.pptx', '**/*.xlsx'). Pipe to grep to search: ox cat '*.docx' | grep -i budget" };
        command.Arguments.Add(fileArg);
        command.Options.Add(sheetOption);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var patterns = parseResult.GetValue(fileArg)!;
            var sheet = parseResult.GetValue(sheetOption);
            var files = GlobExpander.Expand(patterns, GlobExpander.AllFormats);

            if (files.Count == 0)
            {
                Console.Error.WriteLine("error: no matching .docx/.pptx/.xlsx files found");
                return Task.FromResult(1);
            }

            bool multiFile = files.Count > 1;

            foreach (var file in files)
            {
                try
                {
                    var paragraphs = ExtractParagraphs(file, sheet);

                    foreach (var para in paragraphs)
                    {
                        if (multiFile)
                            Console.WriteLine($"{file}:{para.Text}");
                        else
                            Console.WriteLine(para.Text);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"error: {file}: {ex.Message}");
                }
            }

            return Task.FromResult(0);
        });

        return command;
    }

    internal static List<ParagraphInfo> ExtractParagraphs(string file, string? sheet = null)
    {
        var ext = Path.GetExtension(file).ToLowerInvariant();

        return ext switch
        {
            ".docx" => ExtractDocx(file),
            ".pptx" => ExtractPptx(file),
            ".xlsx" => ExtractXlsx(file, sheet),
            _ => throw new InvalidOperationException($"unsupported file format: {ext}")
        };
    }

    private static List<ParagraphInfo> ExtractDocx(string file)
    {
        using var doc = DocumentService.OpenRead(file);
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("invalid .docx: no document body");
        return DocumentService.GetParagraphs(body);
    }

    private static List<ParagraphInfo> ExtractPptx(string file)
    {
        using var doc = PresentationDocument.Open(file, false);
        return PresentationService.GetParagraphs(doc);
    }

    private static List<ParagraphInfo> ExtractXlsx(string file, string? sheet)
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        return SpreadsheetService.GetParagraphs(doc, sheet);
    }
}
