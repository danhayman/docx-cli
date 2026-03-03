namespace Ox.Output;

using Ox.Core;

public static class TextFormatter
{
    public static List<string> FormatParagraphsToLines(IList<ParagraphInfo> paragraphs, int wrap = 80, bool noWrap = false, bool plain = false)
    {
        var outputLines = new List<string>();

        foreach (var para in paragraphs)
        {
            if (plain)
            {
                outputLines.Add(para.Text);
                continue;
            }

            var label = $"P{para.Number}";
            var indent = new string(' ', label.Length + 2);

            if (noWrap || wrap <= 0)
            {
                outputLines.Add($"{label,-4} {para.Text}");
                continue;
            }

            var lines = WrapText(para.Text, wrap - label.Length - 2);
            for (int i = 0; i < lines.Count; i++)
            {
                if (i == 0)
                    outputLines.Add($"{label,-4} {lines[i]}");
                else
                    outputLines.Add($"     {lines[i]}");
            }
        }

        return outputLines;
    }

    public static void WriteParagraphs(IList<ParagraphInfo> paragraphs, int wrap = 80, bool noWrap = false, bool plain = false)
    {
        var lines = FormatParagraphsToLines(paragraphs, wrap, noWrap, plain);
        foreach (var line in lines)
        {
            Console.WriteLine(line);
        }
    }

    public static void WriteParagraphsWithTrackChanges(IList<ParagraphInfo> paragraphs,
        DocumentFormat.OpenXml.Wordprocessing.Body body, int wrap = 80, bool noWrap = false, bool plain = false)
    {
        foreach (var para in paragraphs)
        {
            if (para.Paragraph == null)
            {
                // Table row or synthetic paragraph
                var label = plain ? "" : $"P{para.Number}";
                Console.WriteLine(plain ? para.Text : $"{label,-4} {para.Text}");
                continue;
            }

            var text = BuildTrackChangesText(para.Paragraph);
            var labelStr = $"P{para.Number}";

            if (plain)
            {
                Console.WriteLine(text);
                continue;
            }

            if (noWrap || wrap <= 0)
            {
                Console.WriteLine($"{labelStr,-4} {text}");
                continue;
            }

            var lines = WrapText(text, wrap - labelStr.Length - 2);
            for (int i = 0; i < lines.Count; i++)
            {
                if (i == 0)
                    Console.WriteLine($"{labelStr,-4} {lines[i]}");
                else
                    Console.WriteLine($"     {lines[i]}");
            }
        }
    }

    public static string BuildTrackChangesText(DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
    {
        var sb = new System.Text.StringBuilder();

        foreach (var child in para.ChildElements)
        {
            if (child is DocumentFormat.OpenXml.Wordprocessing.Run run)
            {
                sb.Append(run.InnerText);
            }
            else if (child is DocumentFormat.OpenXml.Wordprocessing.DeletedRun del)
            {
                var text = string.Join("", del.Descendants<DocumentFormat.OpenXml.Wordprocessing.DeletedText>().Select(t => t.Text));
                if (text.Length > 0)
                    sb.Append($"~~{text}~~");
            }
            else if (child is DocumentFormat.OpenXml.Wordprocessing.InsertedRun ins)
            {
                var text = ins.InnerText;
                if (text.Length > 0)
                    sb.Append($"**{text}**");
            }
        }

        return sb.ToString();
    }

    public static List<string> WrapText(string text, int maxWidth)
    {
        if (maxWidth <= 0 || string.IsNullOrEmpty(text))
            return [text];

        if (text.Length <= maxWidth)
            return [text];

        var lines = new List<string>();
        var words = text.Split(' ');
        var currentLine = new System.Text.StringBuilder();

        foreach (var word in words)
        {
            if (currentLine.Length == 0)
            {
                currentLine.Append(word);
            }
            else if (currentLine.Length + 1 + word.Length <= maxWidth)
            {
                currentLine.Append(' ');
                currentLine.Append(word);
            }
            else
            {
                lines.Add(currentLine.ToString());
                currentLine.Clear();
                currentLine.Append(word);
            }
        }

        if (currentLine.Length > 0)
            lines.Add(currentLine.ToString());

        return lines.Count > 0 ? lines : [""];
    }

    public static void WriteTable(string label, IEnumerable<(string Key, string Value)> rows)
    {
        int maxKeyLen = 0;
        var rowList = rows.ToList();
        foreach (var (key, _) in rowList)
        {
            if (key.Length > maxKeyLen) maxKeyLen = key.Length;
        }

        foreach (var (key, value) in rowList)
        {
            Console.WriteLine($"{key.PadRight(maxKeyLen)}  {value}");
        }
    }
}
