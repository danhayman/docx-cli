using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Ox.Core;

public static class TextReplacer
{
    public static int ReplaceInBody(Body body, string oldText, string newText,
        bool replaceAll = false, StringComparison comparison = StringComparison.Ordinal)
    {
        int totalReplaced = 0;

        foreach (var para in body.Descendants<Paragraph>())
        {
            totalReplaced += ReplaceInParagraph(para, oldText, newText, replaceAll, comparison);
            if (totalReplaced > 0 && !replaceAll)
                return totalReplaced;
        }

        return totalReplaced;
    }

    public static int CountInBody(Body body, string searchText, StringComparison comparison = StringComparison.Ordinal)
    {
        int count = 0;
        foreach (var para in body.Descendants<Paragraph>())
        {
            count += TextSearch.FindInParagraph(para, searchText, comparison).Count;
        }
        return count;
    }

    public static int ReplaceInParagraph(Paragraph para, string oldText, string newText,
        bool replaceAll = false, StringComparison comparison = StringComparison.Ordinal)
    {
        var runs = para.Elements<Run>().ToList();
        if (runs.Count == 0) return 0;

        var matches = TextSearch.FindInRuns(runs, oldText, comparison);
        if (matches.Count == 0) return 0;

        // Check if newText contains newlines - if so, we need special handling
        if (newText.Contains('\n'))
        {
            // For newline replacements, only handle the first match for now
            var match = matches[0];
            ApplyReplacementWithNewlines(para, runs, match, newText);
            return 1;
        }

        // Process matches right-to-left so offsets stay valid
        for (int m = matches.Count - 1; m >= 0; m--)
        {
            if (!replaceAll && m < matches.Count - 1)
                break; // Only replace the last match (right-to-left, so this is the first found)

            var match = matches[m];
            ApplyReplacement(runs, match, newText);

            // Rebuild runs list after modification
            runs = para.Elements<Run>().ToList();

            if (!replaceAll) return 1;
        }

        return matches.Count;
    }

    private static void ApplyReplacement(IList<Run> runs, TextMatch match, string newText)
    {
        if (match.StartRunIndex == match.EndRunIndex)
        {
            // Single-run replacement
            ReplaceSingleRun(runs[match.StartRunIndex], match.StartOffsetInRun, match.Length, newText);
        }
        else
        {
            // Cross-run replacement
            ReplaceCrossRun(runs, match, newText);
        }
    }

    private static void ReplaceSingleRun(Run run, int offset, int length, string newText)
    {
        var currentText = run.InnerText;
        var prefix = currentText[..offset];
        var suffix = currentText[(offset + length)..];

        // Remove existing text elements
        foreach (var t in run.Elements<Text>().ToList())
            t.Remove();

        var combined = prefix + newText + suffix;
        if (combined.Length > 0)
        {
            run.Append(new Text(combined) { Space = SpaceProcessingModeValues.Preserve });
        }
    }

    private static void ReplaceCrossRun(IList<Run> runs, TextMatch match, string newText)
    {
        var startRun = runs[match.StartRunIndex];
        var endRun = runs[match.EndRunIndex];

        // Get text parts to keep
        var startText = startRun.InnerText;
        var prefix = startText[..match.StartOffsetInRun];

        var endText = endRun.InnerText;
        var suffix = endText[match.EndOffsetInRun..];

        // Clone RunProperties from the start run for the new text
        var props = startRun.RunProperties?.CloneNode(true) as RunProperties;

        // Create new run with replacement text
        var replacementRun = new Run();
        if (props != null)
            replacementRun.Append(props);
        replacementRun.Append(new Text(prefix + newText + suffix) { Space = SpaceProcessingModeValues.Preserve });

        // Insert new run before start run
        startRun.InsertBeforeSelf(replacementRun);

        // Remove all runs from start to end (inclusive)
        for (int i = match.StartRunIndex; i <= match.EndRunIndex; i++)
        {
            runs[i].Remove();
        }
    }

    private static void ApplyReplacementWithNewlines(Paragraph para, IList<Run> runs, TextMatch match, string newText)
    {
        var startRun = runs[match.StartRunIndex];
        var endRun = runs[match.EndRunIndex];

        // Get text parts to keep
        var startText = startRun.InnerText;
        var prefix = startText[..match.StartOffsetInRun];

        var endText = endRun.InnerText;
        var suffix = endText[match.EndOffsetInRun..];

        // Split replacement text by newlines
        var lines = newText.Split('\n');

        // Clone RunProperties and ParagraphProperties for new paragraphs
        var runProps = startRun.RunProperties?.CloneNode(true) as RunProperties;
        var paraProps = para.ParagraphProperties?.CloneNode(true) as ParagraphProperties;

        // Replace matched text in current paragraph with first line
        var firstLineText = prefix + lines[0] + (lines.Length == 1 ? suffix : "");
        var firstLineRun = new Run();
        if (runProps != null)
            firstLineRun.Append(runProps.CloneNode(true));
        firstLineRun.Append(new Text(firstLineText) { Space = SpaceProcessingModeValues.Preserve });

        startRun.InsertBeforeSelf(firstLineRun);

        // Remove all runs from start to end (inclusive)
        for (int i = match.StartRunIndex; i <= match.EndRunIndex; i++)
        {
            runs[i].Remove();
        }

        // Create new paragraphs for remaining lines
        Paragraph lastPara = para;
        for (int i = 1; i < lines.Length; i++)
        {
            var newPara = new Paragraph();

            // Copy paragraph properties (formatting, bullets, etc.)
            if (paraProps != null)
                newPara.Append(paraProps.CloneNode(true));

            var lineText = i == lines.Length - 1 ? lines[i] + suffix : lines[i];
            var newRun = new Run();
            if (runProps != null)
                newRun.Append(runProps.CloneNode(true));
            newRun.Append(new Text(lineText) { Space = SpaceProcessingModeValues.Preserve });
            newPara.Append(newRun);

            lastPara.InsertAfterSelf(newPara);
            lastPara = newPara;
        }
    }
}
