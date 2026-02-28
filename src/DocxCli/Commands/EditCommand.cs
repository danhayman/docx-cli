using System.CommandLine;
using DocxCli.Core;
namespace DocxCli.Commands;

public static class EditCommand
{
    public static Command Create()
    {
        var fileArg = new Argument<string[]>("file") { Description = "Path(s) or glob pattern(s) for .docx files", Arity = ArgumentArity.OneOrMore };
        var oldOption = new Option<string>("--old") { Description = "Text to find", Required = true };
        var newOption = new Option<string>("--new") { Description = "Replacement text", Required = true };
        var replaceAllOption = new Option<bool>("--replace-all") { Description = "Replace all occurrences" };
        var trackOption = new Option<bool>("--track") { Description = "Insert as tracked change" };
        var authorOption = new Option<string>("--author") { Description = "Author for tracked changes", DefaultValueFactory = _ => "docx-cli" };
        var backupOption = new Option<bool>("--backup") { Description = "Create .docx.bak before editing" };
        var outputOption = new Option<string?>("--output", "-o") { Description = "Write to different file instead of in-place" };

        var command = new Command("edit") { Description = "Replace text in documents (use \\n for new paragraphs)" };
        command.Arguments.Add(fileArg);
        command.Options.Add(oldOption);
        command.Options.Add(newOption);
        command.Options.Add(replaceAllOption);
        command.Options.Add(trackOption);
        command.Options.Add(authorOption);
        command.Options.Add(backupOption);
        command.Options.Add(outputOption);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var patterns = parseResult.GetValue(fileArg)!;
            var oldText = UnescapeString(parseResult.GetValue(oldOption)!);
            var newText = UnescapeString(parseResult.GetValue(newOption)!);
            var replaceAll = parseResult.GetValue(replaceAllOption);
            var track = parseResult.GetValue(trackOption);
            var author = parseResult.GetValue(authorOption)!;
            var backup = parseResult.GetValue(backupOption);
            var output = parseResult.GetValue(outputOption);
            var dryRun = parseResult.GetValue(GlobalOptions.DryRun);
            var force = parseResult.GetValue(GlobalOptions.Force);
            var files = GlobExpander.Expand(patterns);

            if (files.Count == 0)
            {
                Console.Error.WriteLine("error: no matching .docx files found");
                return Task.FromResult(1);
            }

            // --output only valid for single file
            if (output != null && files.Count > 1)
            {
                Console.Error.WriteLine("error: --output cannot be used with multiple files");
                return Task.FromResult(1);
            }

            bool multiFile = files.Count > 1;
            int totalReplaced = 0;
            int filesChanged = 0;

            foreach (var file in files)
            {
                try
                {
                    // Lock check
                    if (!dryRun)
                        LockDetector.CheckLock(file, force);

                    if (dryRun)
                    {
                        using var readDoc = DocumentService.OpenRead(file);
                        var readBody = readDoc.MainDocumentPart!.Document.Body!;
                        var count = TextReplacer.CountInBody(readBody, oldText);

                        if (count == 0)
                        {
                            if (!multiFile)
                            {
                                Console.Error.WriteLine($"error: \"{oldText}\" not found");
                                return Task.FromResult(1);
                            }
                            continue;
                        }

                        if (count > 1 && !replaceAll)
                        {
                            Console.Error.WriteLine($"error: {(multiFile ? file + ": " : "")}\"{oldText}\" is not unique (found {count} occurrences). Use --replace-all to replace all.");
                            if (!multiFile)
                                return Task.FromResult(1);
                            continue;
                        }

                        var msg = multiFile
                            ? $"{file}: would replace {count} occurrence{(count != 1 ? "s" : "")}"
                            : $"would replace {count} occurrence{(count != 1 ? "s" : "")}";
                        Console.WriteLine(msg);
                        totalReplaced += count;
                        continue;
                    }

                    // Pre-check
                    {
                        using var checkDoc = DocumentService.OpenRead(file);
                        var checkBody = checkDoc.MainDocumentPart!.Document.Body!;
                        var count = TextReplacer.CountInBody(checkBody, oldText);

                        if (count == 0)
                        {
                            if (!multiFile)
                            {
                                Console.Error.WriteLine($"error: \"{oldText}\" not found");
                                return Task.FromResult(1);
                            }
                            continue;
                        }

                        if (count > 1 && !replaceAll)
                        {
                            Console.Error.WriteLine($"error: {(multiFile ? file + ": " : "")}\"{oldText}\" is not unique (found {count} occurrences). Use --replace-all to replace all.");
                            if (!multiFile)
                                return Task.FromResult(1);
                            continue;
                        }
                    }

                    using var doc = DocumentService.OpenForEdit(file, output, backup);
                    var body = doc.MainDocumentPart!.Document.Body!;

                    int replaced;
                    if (track)
                    {
                        replaced = TrackedChanges.ReplaceWithTracking(body, oldText, newText, author, replaceAll);
                        DocumentService.SaveAtomically(doc, file, output);
                    }
                    else
                    {
                        replaced = TextReplacer.ReplaceInBody(body, oldText, newText, replaceAll);
                        DocumentService.SaveAtomically(doc, file, output);
                    }

                    totalReplaced += replaced;
                    filesChanged++;

                    if (multiFile)
                        Console.WriteLine($"{file}: replaced {replaced} occurrence{(replaced != 1 ? "s" : "")}");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"error: {file}: {ex.Message}");
                }
            }

            if (!dryRun && !multiFile && filesChanged > 0)
            {
                Console.WriteLine($"replaced {totalReplaced} occurrence{(totalReplaced != 1 ? "s" : "")}");
            }
            else if (multiFile && filesChanged > 0)
            {
                Console.WriteLine($"total: {totalReplaced} replacement{(totalReplaced != 1 ? "s" : "")} across {filesChanged} file{(filesChanged != 1 ? "s" : "")}");
            }
            else if (filesChanged == 0 && !dryRun)
            {
                Console.Error.WriteLine($"error: \"{oldText}\" not found in any file");
                return Task.FromResult(1);
            }

            return Task.FromResult(0);
        });

        return command;
    }

    private static string UnescapeString(string s)
    {
        if (!s.Contains('\\'))
            return s;

        var result = new System.Text.StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '\\' && i + 1 < s.Length)
            {
                switch (s[i + 1])
                {
                    case 'n':
                        result.Append('\n');
                        i++;
                        break;
                    case 't':
                        result.Append('\t');
                        i++;
                        break;
                    case '\\':
                        result.Append('\\');
                        i++;
                        break;
                    default:
                        result.Append(s[i]);
                        break;
                }
            }
            else
            {
                result.Append(s[i]);
            }
        }
        return result.ToString();
    }
}
