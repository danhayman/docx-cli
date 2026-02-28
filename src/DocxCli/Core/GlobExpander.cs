namespace DocxCli.Core;

public static class GlobExpander
{
    public static List<string> Expand(string[] patterns)
    {
        var results = new List<string>();
        foreach (var pattern in patterns)
        {
            if (ContainsGlobChars(pattern))
                results.AddRange(ExpandGlob(pattern));
            else
                results.Add(pattern);
        }
        return results;
    }

    private static bool ContainsGlobChars(string pattern)
    {
        return pattern.Contains('*') || pattern.Contains('?');
    }

    private static IEnumerable<string> ExpandGlob(string pattern)
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
        return Directory.EnumerateFiles(baseDir, searchPattern, option)
            .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase);
    }
}
