using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Ox.Core;

public static class TrackedChanges
{
    public static int ReplaceWithTracking(Body body, string oldText, string newText,
        string author, bool replaceAll = false, StringComparison comparison = StringComparison.Ordinal)
    {
        int totalReplaced = 0;
        var revId = GetNextRevisionId(body);

        foreach (var para in body.Descendants<Paragraph>())
        {
            var runs = para.Elements<Run>().ToList();
            if (runs.Count == 0) continue;

            var matches = TextSearch.FindInRuns(runs, oldText, comparison);
            if (matches.Count == 0) continue;

            // Process right-to-left
            for (int m = matches.Count - 1; m >= 0; m--)
            {
                if (!replaceAll && totalReplaced > 0) break;

                var match = matches[m];
                ApplyTrackedReplacement(runs, match, oldText, newText, author, revId.ToString(), (revId + 1).ToString());
                revId += 2;
                totalReplaced++;

                // Rebuild runs after modification
                runs = para.Elements<Run>().ToList();

                if (!replaceAll) return totalReplaced;
            }
        }

        return totalReplaced;
    }

    public static int InsertWithTracking(Paragraph para, string text, string author, int revId)
    {
        var run = new Run();
        var existingProps = para.Elements<Run>().FirstOrDefault()?.RunProperties;
        if (existingProps != null)
            run.Append(existingProps.CloneNode(true));
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        var ins = new InsertedRun
        {
            Author = author,
            Date = DateTime.UtcNow,
            Id = revId.ToString()
        };
        ins.Append(run);
        para.Append(ins);

        return 1;
    }

    private static void ApplyTrackedReplacement(IList<Run> runs, TextMatch match,
        string oldText, string newText, string author, string delRevId, string insRevId)
    {
        var now = DateTime.UtcNow;
        var startRun = runs[match.StartRunIndex];

        // Clone formatting from start run
        var props = startRun.RunProperties?.CloneNode(true) as RunProperties;

        // Build the deleted run
        var deletedRun = new Run();
        if (props != null)
            deletedRun.Append(props.CloneNode(true) as RunProperties ?? new RunProperties());
        deletedRun.Append(new DeletedText(oldText) { Space = SpaceProcessingModeValues.Preserve });

        var del = new DeletedRun
        {
            Author = author,
            Date = now,
            Id = delRevId
        };
        del.Append(deletedRun);

        // Build the inserted run
        var insertedRun = new Run();
        if (props != null)
            insertedRun.Append(props.CloneNode(true) as RunProperties ?? new RunProperties());
        insertedRun.Append(new Text(newText) { Space = SpaceProcessingModeValues.Preserve });

        var ins = new InsertedRun
        {
            Author = author,
            Date = now,
            Id = insRevId
        };
        ins.Append(insertedRun);

        // Handle prefix/suffix text in boundary runs
        var startText = startRun.InnerText;
        var prefix = startText[..match.StartOffsetInRun];

        var endRun = runs[match.EndRunIndex];
        var endText = endRun.InnerText;
        var suffix = endText[match.EndOffsetInRun..];

        // Insert tracked changes before the start run
        var insertPoint = startRun;

        if (prefix.Length > 0)
        {
            var prefixRun = new Run();
            if (props != null)
                prefixRun.Append(props.CloneNode(true) as RunProperties ?? new RunProperties());
            prefixRun.Append(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve });
            insertPoint.InsertBeforeSelf(prefixRun);
        }

        insertPoint.InsertBeforeSelf(del);
        insertPoint.InsertBeforeSelf(ins);

        if (suffix.Length > 0)
        {
            var suffixRun = new Run();
            if (props != null)
                suffixRun.Append(props.CloneNode(true) as RunProperties ?? new RunProperties());
            suffixRun.Append(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve });
            insertPoint.InsertBeforeSelf(suffixRun);
        }

        // Remove original runs
        for (int i = match.StartRunIndex; i <= match.EndRunIndex; i++)
        {
            runs[i].Remove();
        }
    }

    public static int GetNextRevisionId(Body body)
    {
        int maxId = 0;

        foreach (var del in body.Descendants<DeletedRun>())
        {
            if (del.Id?.Value != null && int.TryParse(del.Id.Value, out var id) && id > maxId)
                maxId = id;
        }

        foreach (var ins in body.Descendants<InsertedRun>())
        {
            if (ins.Id?.Value != null && int.TryParse(ins.Id.Value, out var id) && id > maxId)
                maxId = id;
        }

        return maxId + 1;
    }

    public static List<TrackedChange> GetTrackedChanges(Body body)
    {
        var changes = new List<TrackedChange>();

        foreach (var del in body.Descendants<DeletedRun>())
        {
            var text = string.Join("", del.Descendants<DeletedText>().Select(t => t.Text));
            changes.Add(new TrackedChange("deleted", text, del.Author?.Value, del.Date?.Value));
        }

        foreach (var ins in body.Descendants<InsertedRun>())
        {
            var text = ins.InnerText;
            changes.Add(new TrackedChange("inserted", text, ins.Author?.Value, ins.Date?.Value));
        }

        return changes;
    }
}

public record TrackedChange(string Type, string Text, string? Author, DateTimeOffset? Date);
