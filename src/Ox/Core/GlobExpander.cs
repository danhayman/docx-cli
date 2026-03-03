namespace Ox.Core;

public static class GlobExpander
{
    public static readonly string[] DocxOnly = [".docx"];
    public static readonly string[] AllFormats = [".docx", ".pptx", ".xlsx"];

    public static List<string> Expand(string[] patterns, string[]? extensions = null)
    {
        var results = new List<string>();
        foreach (var pattern in patterns)
        {
            if (ContainsGlobChars(pattern))
                results.AddRange(ExpandGlob(pattern, extensions));
            else
                results.Add(pattern);
        }
        return results;
    }

    private static bool ContainsGlobChars(string pattern)
    {
        return pattern.Contains('*') || pattern.Contains('?');
    }

    private static IEnumerable<string> ExpandGlob(string pattern, string[]? extensions)
    {
        // Normalize separators
        pattern = pattern.Replace('\\', '/');

        // Split into directory part and file pattern
        string baseDir;
        string searchPattern;
        bool recursive = pattern.Contains("**/");

        if (recursive)
        {
            // "**/*.docx" or "dir/**/*.docx"
            var parts = pattern.Split("**/", 2);
            baseDir = string.IsNullOrEmpty(parts[0]) ? "." : parts[0].TrimEnd('/');
            searchPattern = parts[1].TrimStart('/');
        }
        else
        {
            var lastSlash = pattern.LastIndexOf('/');
            if (lastSlash >= 0)
            {
                baseDir = pattern[..lastSlash];
                searchPattern = pattern[(lastSlash + 1)..];
            }
            else
            {
                baseDir = ".";
                searchPattern = pattern;
            }
        }

        if (!Directory.Exists(baseDir))
            return [];

        var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var files = Directory.EnumerateFiles(baseDir, searchPattern, option);
        if (extensions != null)
            files = files.Where(f => extensions.Any(ext => f.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));
        return files.OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
    }
}
