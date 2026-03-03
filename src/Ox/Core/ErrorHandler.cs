using System.CommandLine;
using DocumentFormat.OpenXml.Packaging;

namespace Ox.Core;

public static class ErrorHandler
{
    public static void SetActionWithErrorHandling(this Command command, Func<ParseResult, CancellationToken, Task<int>> handler)
    {
        command.SetAction(async (parseResult, ct) =>
        {
            try
            {
                return await handler(parseResult, ct);
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"error: file not found: {ex.FileName}");
                return 1;
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.Error.WriteLine($"error: directory not found: {ex.Message}");
                return 1;
            }
            catch (IOException ex) when (ex.Message.Contains("being used by another process"))
            {
                Console.Error.WriteLine("error: file is locked (close Word/LibreOffice or use --force)");
                return 1;
            }
            catch (IOException ex)
            {
                Console.Error.WriteLine($"error: I/O error: {ex.Message}");
                return 1;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"error: access denied: {ex.Message}");
                return 1;
            }
            catch (OpenXmlPackageException ex)
            {
                Console.Error.WriteLine($"error: invalid or corrupted .docx file: {ex.Message}");
                return 1;
            }
            catch (InvalidOperationException ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                return 1;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                return 1;
            }
        });
    }
}
