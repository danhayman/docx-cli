using System.CommandLine;
using Ox.Core;
using Ox.Output;

namespace Ox.Commands;

public static class CommentCommand
{
    public static Command Create()
    {
        var command = new Command("comment") { Description = "Add, list, or remove comments" };

        command.Subcommands.Add(CreateAddCommand());
        command.Subcommands.Add(CreateListCommand());
        command.Subcommands.Add(CreateDeleteCommand());

        return command;
    }

    private static Command CreateAddCommand()
    {
        var fileArg = new Argument<string>("file") { Description = "Path to .docx file" };
        var atOption = new Option<string>("--at") { Description = "Anchor text for the comment", Required = true };
        var textOption = new Option<string>("--text") { Description = "Comment text", Required = true };
        var authorOption = new Option<string>("--author") { Description = "Comment author", DefaultValueFactory = _ => "ox" };
        var backupOption = new Option<bool>("--backup") { Description = "Create .docx.bak before editing" };
        var outputOption = new Option<string?>("--output", "-o") { Description = "Write to different file" };

        var command = new Command("add") { Description = "Add a comment anchored to text" };
        command.Arguments.Add(fileArg);
        command.Options.Add(atOption);
        command.Options.Add(textOption);
        command.Options.Add(authorOption);
        command.Options.Add(backupOption);
        command.Options.Add(outputOption);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var file = parseResult.GetValue(fileArg)!;
            var at = parseResult.GetValue(atOption)!;
            var text = parseResult.GetValue(textOption)!;
            var author = parseResult.GetValue(authorOption)!;
            var backup = parseResult.GetValue(backupOption);
            var output = parseResult.GetValue(outputOption);
            var force = parseResult.GetValue(GlobalOptions.Force);

            LockDetector.CheckLock(file, force);

            using var doc = DocumentService.OpenForEdit(file, output, backup);
            var commentId = CommentService.AddComment(doc, at, text, author);
            DocumentService.SaveAtomically(doc, file, output);

            Console.WriteLine($"comment added at \"{at}\" (id: {commentId})");

            return Task.FromResult(0);
        });

        return command;
    }

    private static Command CreateListCommand()
    {
        var fileArg = new Argument<string>("file") { Description = "Path to .docx file" };

        var command = new Command("list") { Description = "List all comments" };
        command.Arguments.Add(fileArg);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var file = parseResult.GetValue(fileArg)!;

            using var doc = DocumentService.OpenRead(file);
            var comments = CommentService.ListComments(doc);

            if (comments.Count == 0)
            {
                Console.WriteLine("No comments.");
                return Task.FromResult(0);
            }

            Console.WriteLine($"{"ID",-4} {"AUTHOR",-12} {"DATE",-12} {"AT",-20} TEXT");
            foreach (var c in comments)
            {
                var dateStr = c.Date?.ToString("yyyy-MM-dd") ?? "";
                var anchor = c.AnchorText.Length > 18 ? c.AnchorText[..18] + ".." : c.AnchorText;
                Console.WriteLine($"{c.Id,-4} {c.Author,-12} {dateStr,-12} \"{anchor}\"".PadRight(50) + c.Text);
            }

            return Task.FromResult(0);
        });

        return command;
    }

    private static Command CreateDeleteCommand()
    {
        var fileArg = new Argument<string>("file") { Description = "Path to .docx file" };
        var idOption = new Option<int>("--id") { Description = "Comment ID to delete", Required = true };
        var backupOption = new Option<bool>("--backup") { Description = "Create .docx.bak before editing" };
        var outputOption = new Option<string?>("--output", "-o") { Description = "Write to different file" };

        var command = new Command("delete") { Description = "Delete a comment by ID" };
        command.Arguments.Add(fileArg);
        command.Options.Add(idOption);
        command.Options.Add(backupOption);
        command.Options.Add(outputOption);

        command.SetActionWithErrorHandling((parseResult, ct) =>
        {
            var file = parseResult.GetValue(fileArg)!;
            var id = parseResult.GetValue(idOption);
            var backup = parseResult.GetValue(backupOption);
            var output = parseResult.GetValue(outputOption);
            var force = parseResult.GetValue(GlobalOptions.Force);

            LockDetector.CheckLock(file, force);

            using var doc = DocumentService.OpenForEdit(file, output, backup);
            var deleted = CommentService.DeleteComment(doc, id);
            DocumentService.SaveAtomically(doc, file, output);

            if (deleted)
                Console.WriteLine($"comment {id} deleted");
            else
            {
                Console.Error.WriteLine($"error: comment {id} not found");
                return Task.FromResult(1);
            }

            return Task.FromResult(0);
        });

        return command;
    }
}
