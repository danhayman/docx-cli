using System.CommandLine;
using Ox;
using Ox.Commands;

var rootCommand = new RootCommand("Read and edit Office documents (.docx, .pptx, .xlsx) from the terminal.");

rootCommand.Add(GlobalOptions.DryRun);
rootCommand.Add(GlobalOptions.Force);

rootCommand.Subcommands.Add(CatCommand.Create());
rootCommand.Subcommands.Add(ReadCommand.Create());
rootCommand.Subcommands.Add(InfoCommand.Create());
rootCommand.Subcommands.Add(EditCommand.Create());
rootCommand.Subcommands.Add(CommentCommand.Create());

return await rootCommand.Parse(args).InvokeAsync();
