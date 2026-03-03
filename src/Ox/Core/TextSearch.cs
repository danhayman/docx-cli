using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Ox.Core;

public record RunMapping(int RunIndex, int StartInConcat, int Length);

public record TextMatch(
    int StartInConcat,
    int Length,
    int StartRunIndex,
    int StartOffsetInRun,
    int EndRunIndex,
    int EndOffsetInRun);

public static class TextSearch
{
    public static (string FullText, List<RunMapping> RunMap) BuildRunMap(IList<Run> runs)
    {
        var sb = new StringBuilder();
        var runMap = new List<RunMapping>();

        for (int i = 0; i < runs.Count; i++)
        {
            var text = runs[i].InnerText;
            runMap.Add(new RunMapping(i, sb.Length, text.Length));
            sb.Append(text);
        }

        return (sb.ToString(), runMap);
    }

    public static List<TextMatch> FindInParagraph(Paragraph para, string searchText, StringComparison comparison = StringComparison.Ordinal)
    {
        var runs = para.Elements<Run>().ToList();
        return FindInRuns(runs, searchText, comparison);
    }

    public static List<TextMatch> FindInRuns(IList<Run> runs, string searchText, StringComparison comparison = StringComparison.Ordinal)
    {
        if (runs.Count == 0 || string.IsNullOrEmpty(searchText))
            return [];

        var (fullText, runMap) = BuildRunMap(runs);
        var matches = new List<TextMatch>();

        int pos = 0;
        while (pos <= fullText.Length - searchText.Length)
        {
            int idx = fullText.IndexOf(searchText, pos, comparison);
            if (idx < 0) break;

            var (startRun, startOffset) = MapOffsetToRun(runMap, idx);
            var (endRun, endOffset) = MapOffsetToRun(runMap, idx + searchText.Length);

            matches.Add(new TextMatch(idx, searchText.Length, startRun, startOffset, endRun, endOffset));
            pos = idx + searchText.Length;
        }

        return matches;
    }

    internal static (int RunIndex, int OffsetInRun) MapOffsetToRun(List<RunMapping> runMap, int offset)
    {
        for (int i = runMap.Count - 1; i >= 0; i--)
        {
            if (offset >= runMap[i].StartInConcat)
            {
                return (runMap[i].RunIndex, offset - runMap[i].StartInConcat);
            }
        }

        return (0, offset);
    }
}
