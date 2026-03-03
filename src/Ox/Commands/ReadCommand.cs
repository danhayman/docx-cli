using System.CommandLine;
using Ox.Core;
using Ox.Output;

namespace Ox.Commands;

public static class ReadCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string[]>("file") { Description = "Path(s) or glob pattern(s) for .docx, .pptx, or .xlsx files", Arity = ArgumentArity.OneOrMore };
        var offsetOption = new Option<int>("--offset") { Description = "Character offset to start from (0-indexed)" };
        var limitOption = new Option<int>("--limit") { Description = "Max characters to read (0 = all)", DefaultValueFactory = _ => 100000 };
        var trackChangesOption = new Option<bool>("--track-changes") { Description = "Show tracked changes inline (.docx only)" };
        var sheetOption = new Option<string?>("--sheet") { Description = "For .xlsx: only output rows from this sheet (default: all sheets)" };

        var command = new Command("read") { Description = "Read document content with character offset/limit. Supports globs (e.g. '*.docx', '**/*.pptx')" };
        command.Arguments.Add(fileArg);
        command.Options.Add(offsetOption);
        command.Options.Add(limitOption);
        command.Options.Add(trackChangesOption);
        command.Options.Add(sheetOption);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var patterns = parseResult.GetValue(fileArg)!;
            var offset = parseResult.GetValue(offsetOption);
            var limit = parseResult.GetValue(limitOption);
            var trackChanges = parseResult.GetValue(trackChangesOption);
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
                    var paragraphs = CatCommand.ExtractParagraphs(file, sheet);

                    var output = new System.Text.StringBuilder();
                    foreach (var para in paragraphs)
                    {
                        if (trackChanges && para.Paragraph != null)
                        {
                            var text = TextFormatter.BuildTrackChangesText(para.Paragraph);
                            output.AppendLine(text);
                        }
                        else
                        {
                            output.AppendLine(para.Text);
                        }
                    }

                    var fullText = output.ToString();

                    // Apply character-based offset and limit
                    if (offset > 0 && offset < fullText.Length)
                        fullText = fullText.Substring(offset);
                    else if (offset >= fullText.Length)
                        fullText = "";

                    if (limit > 0 && fullText.Length > limit)
                        fullText = fullText.Substring(0, limit);

                    if (multiFile)
                    {
                        foreach (var line in fullText.TrimEnd('\n', '\r').Split('\n'))
                            Console.WriteLine($"{file}:{line}");
                    }
                    else
                    {
                        Console.Write(fullText);
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
}
