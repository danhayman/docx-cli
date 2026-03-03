using System.CommandLine;
using DocxCli.Core;

namespace DocxCli.Commands;

public static class CatCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string[]>("file") { Description = "Path(s) or glob pattern(s) for .docx files", Arity = ArgumentArity.OneOrMore };
        var command = new Command("cat") { Description = "Dump plain text. Supports globs (e.g. '*.docx', '**/*.docx'). Pipe to grep to search: docx cat '*.docx' | grep -i budget" };
        command.Arguments.Add(fileArg);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var patterns = parseResult.GetValue(fileArg)!;
            var files = GlobExpander.Expand(patterns);

            if (files.Count == 0)
            {
                Console.Error.WriteLine("error: no matching .docx files found");
                return Task.FromResult(1);
            }

            bool multiFile = files.Count > 1;

            foreach (var file in files)
            {
                try
                {
                    using var doc = DocumentService.OpenRead(file);
                    var body = doc.MainDocumentPart!.Document.Body!;
                    var paragraphs = DocumentService.GetParagraphs(body);

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
}
