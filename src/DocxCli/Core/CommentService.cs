using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxCli.Core;

public static class CommentService
{
    public static int AddComment(WordprocessingDocument doc, string anchorText, string commentText,
        string author, StringComparison comparison = StringComparison.Ordinal)
    {
        var body = doc.MainDocumentPart!.Document.Body!;

        // Find the anchor text
        foreach (var para in body.Descendants<Paragraph>())
        {
            var runs = para.Elements<Run>().ToList();
            var matches = TextSearch.FindInRuns(runs, anchorText, comparison);
            if (matches.Count == 0) continue;

            var match = matches[0];
            var commentId = GetNextCommentId(doc);

            // Ensure comments part exists
            var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart
                ?? doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments ??= new Comments();

            // Create the comment
            var comment = new Comment
            {
                Id = commentId.ToString(),
                Author = author,
                Date = DateTime.UtcNow,
                Initials = author.Length > 0 ? author[0].ToString() : "A"
            };
            comment.Append(new Paragraph(new Run(new Text(commentText))));
            commentsPart.Comments.Append(comment);

            // Split runs at match boundaries to get precise anchoring
            var idStr = commentId.ToString();
            var startRun = runs[match.StartRunIndex];
            var endRun = runs[match.EndRunIndex];

            // Split start run if match doesn't start at beginning
            if (match.StartOffsetInRun > 0)
            {
                var text = startRun.InnerText;
                var prefix = text[..match.StartOffsetInRun];
                var rest = text[match.StartOffsetInRun..];

                // Update start run to contain only prefix
                foreach (var t in startRun.Elements<Text>().ToList()) t.Remove();
                startRun.Append(new Text(prefix) { Space = SpaceProcessingModeValues.Preserve });

                // Create new run for the matched portion
                var matchRun = new Run();
                if (startRun.RunProperties != null)
                    matchRun.Append(startRun.RunProperties.CloneNode(true));
                matchRun.Append(new Text(rest) { Space = SpaceProcessingModeValues.Preserve });
                startRun.InsertAfterSelf(matchRun);

                // Update reference: the anchor now starts at the new run
                startRun = matchRun;
                // If start and end were the same run, update endRun too
                if (match.StartRunIndex == match.EndRunIndex)
                {
                    endRun = matchRun;
                }
            }

            // Split end run if match doesn't end at the end of the run
            var endText = endRun.InnerText;
            var endOffset = match.StartRunIndex == match.EndRunIndex && match.StartOffsetInRun > 0
                ? match.EndOffsetInRun - match.StartOffsetInRun
                : match.EndOffsetInRun;
            if (endOffset < endText.Length)
            {
                var keep = endText[..endOffset];
                var suffix = endText[endOffset..];

                foreach (var t in endRun.Elements<Text>().ToList()) t.Remove();
                endRun.Append(new Text(keep) { Space = SpaceProcessingModeValues.Preserve });

                var suffixRun = new Run();
                if (endRun.RunProperties != null)
                    suffixRun.Append(endRun.RunProperties.CloneNode(true));
                suffixRun.Append(new Text(suffix) { Space = SpaceProcessingModeValues.Preserve });
                endRun.InsertAfterSelf(suffixRun);
            }

            // Now place comment range markers precisely around the anchor runs
            var rangeStart = new CommentRangeStart { Id = idStr };
            var rangeEnd = new CommentRangeEnd { Id = idStr };

            startRun.InsertBeforeSelf(rangeStart);
            endRun.InsertAfterSelf(rangeEnd);

            var refRun = new Run(new CommentReference { Id = idStr });
            rangeEnd.InsertAfterSelf(refRun);

            return commentId;
        }

        throw new InvalidOperationException($"Text \"{anchorText}\" not found in document");
    }

    public static List<CommentInfo> ListComments(WordprocessingDocument doc)
    {
        var commentsPart = doc.MainDocumentPart?.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null)
            return [];

        var result = new List<CommentInfo>();
        foreach (var comment in commentsPart.Comments.Elements<Comment>())
        {
            var id = int.Parse(comment.Id!.Value!);
            var text = comment.InnerText;
            var author = comment.Author?.Value ?? "";
            var date = comment.Date?.Value;

            // Find the anchor text from the document body
            var anchorText = FindCommentAnchorText(doc, comment.Id!.Value!);

            result.Add(new CommentInfo(id, author, date, anchorText, text));
        }

        return result;
    }

    public static bool DeleteComment(WordprocessingDocument doc, int commentId)
    {
        var commentsPart = doc.MainDocumentPart?.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null)
            return false;

        var idStr = commentId.ToString();

        // Remove the comment element
        var comment = commentsPart.Comments.Elements<Comment>()
            .FirstOrDefault(c => c.Id?.Value == idStr);
        if (comment == null) return false;
        comment.Remove();

        // Remove range markers and reference from document body
        var body = doc.MainDocumentPart!.Document.Body!;
        foreach (var start in body.Descendants<CommentRangeStart>().Where(s => s.Id?.Value == idStr).ToList())
            start.Remove();
        foreach (var end in body.Descendants<CommentRangeEnd>().Where(e => e.Id?.Value == idStr).ToList())
            end.Remove();
        foreach (var refEl in body.Descendants<CommentReference>().Where(r => r.Id?.Value == idStr).ToList())
            refEl.Parent?.Remove(); // Remove the containing Run

        return true;
    }

    private static string FindCommentAnchorText(WordprocessingDocument doc, string commentId)
    {
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return "";

        var rangeStart = body.Descendants<CommentRangeStart>()
            .FirstOrDefault(s => s.Id?.Value == commentId);
        var rangeEnd = body.Descendants<CommentRangeEnd>()
            .FirstOrDefault(e => e.Id?.Value == commentId);

        if (rangeStart == null || rangeEnd == null) return "";

        // Collect text between range start and end
        var collecting = false;
        var text = new System.Text.StringBuilder();
        var parent = rangeStart.Parent;
        if (parent == null) return "";

        foreach (var el in parent.ChildElements)
        {
            if (el == rangeStart) { collecting = true; continue; }
            if (el == rangeEnd) break;
            if (collecting && el is Run run)
            {
                text.Append(run.InnerText);
            }
        }

        return text.ToString();
    }

    private static int GetNextCommentId(WordprocessingDocument doc)
    {
        var commentsPart = doc.MainDocumentPart?.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null) return 1;

        int maxId = 0;
        foreach (var comment in commentsPart.Comments.Elements<Comment>())
        {
            if (comment.Id?.Value != null && int.TryParse(comment.Id.Value, out var id) && id > maxId)
                maxId = id;
        }

        return maxId + 1;
    }
}

public record CommentInfo(int Id, string Author, DateTimeOffset? Date, string AnchorText, string Text);
