using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ox.Core;

namespace Ox.Tests;

public class CommentServiceTests
{
    [Fact]
    public void AddAndListComment()
    {
        var path = TestHelper.CreateTestDocx("Hello World, this is a test.");
        using var doc = WordprocessingDocument.Open(path, true);

        var commentId = CommentService.AddComment(doc, "World", "Check this word", "TestAuthor");
        doc.Save();

        Assert.Equal(1, commentId);

        var comments = CommentService.ListComments(doc);
        Assert.Single(comments);
        Assert.Equal("Check this word", comments[0].Text);
        Assert.Equal("TestAuthor", comments[0].Author);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void DeleteComment()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);

        var commentId = CommentService.AddComment(doc, "World", "Comment text", "Author");
        doc.Save();

        var deleted = CommentService.DeleteComment(doc, commentId);
        doc.Save();

        Assert.True(deleted);

        var comments = CommentService.ListComments(doc);
        Assert.Empty(comments);

        // Verify range markers are removed
        var body = doc.MainDocumentPart!.Document.Body!;
        Assert.Empty(body.Descendants<CommentRangeStart>());
        Assert.Empty(body.Descendants<CommentRangeEnd>());

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void AddCommentTextNotFound()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);

        Assert.Throws<InvalidOperationException>(() =>
            CommentService.AddComment(doc, "xyz", "Comment", "Author"));

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void DeleteNonExistentComment()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);

        var deleted = CommentService.DeleteComment(doc, 999);
        Assert.False(deleted);

        doc.Dispose();
        File.Delete(path);
    }
}
