using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxCli.Core;

namespace DocxCli.Tests;

public class TrackedChangesTests
{
    [Fact]
    public void ReplaceWithTrackingSingleRun()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TrackedChanges.ReplaceWithTracking(body, "World", "Earth", "TestAuthor");
        doc.Save();

        Assert.Equal(1, count);

        // Verify tracked change elements exist
        var deletions = body.Descendants<DeletedRun>().ToList();
        var insertions = body.Descendants<InsertedRun>().ToList();

        Assert.Single(deletions);
        Assert.Single(insertions);
        Assert.Equal("TestAuthor", deletions[0].Author?.Value);
        Assert.Equal("TestAuthor", insertions[0].Author?.Value);

        // Verify the deleted text
        var delText = string.Join("", deletions[0].Descendants<DeletedText>().Select(t => t.Text));
        Assert.Equal("World", delText);

        // Verify the inserted text
        Assert.Equal("Earth", insertions[0].InnerText);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void GetTrackedChanges()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        TrackedChanges.ReplaceWithTracking(body, "World", "Earth", "Author1");
        doc.Save();

        var changes = TrackedChanges.GetTrackedChanges(body);
        Assert.Equal(2, changes.Count);

        var deleted = changes.First(c => c.Type == "deleted");
        var inserted = changes.First(c => c.Type == "inserted");

        Assert.Equal("World", deleted.Text);
        Assert.Equal("Earth", inserted.Text);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void RevisionIdAutoIncrements()
    {
        var path = TestHelper.CreateTestDocx("cat and cat");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        TrackedChanges.ReplaceWithTracking(body, "cat", "dog", "Author1", replaceAll: true);
        doc.Save();

        // Each replacement should get unique IDs
        var delIds = body.Descendants<DeletedRun>().Select(d => d.Id?.Value).ToList();
        var insIds = body.Descendants<InsertedRun>().Select(i => i.Id?.Value).ToList();

        var allIds = delIds.Concat(insIds).Where(id => id != null).ToList();
        Assert.Equal(allIds.Count, allIds.Distinct().Count()); // All unique

        doc.Dispose();
        File.Delete(path);
    }
}
