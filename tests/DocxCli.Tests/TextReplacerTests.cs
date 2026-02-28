using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxCli.Core;

namespace DocxCli.Tests;

public class TextReplacerTests
{
    [Fact]
    public void ReplaceSingleRunSimple()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "World", "Earth");
        doc.Save();
        doc.Dispose();

        Assert.Equal(1, count);
        var text = TestHelper.ReadDocxText(path);
        Assert.Equal("Hello Earth", text);

        File.Delete(path);
    }

    [Fact]
    public void ReplaceCrossRunTwoRuns()
    {
        var path = TestHelper.CreateMultiRunDocx(["Hello Wo", "rld!"]);
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "World", "Earth");
        doc.Save();
        doc.Dispose();

        Assert.Equal(1, count);
        var text = TestHelper.ReadDocxText(path);
        Assert.Equal("Hello Earth!", text);

        File.Delete(path);
    }

    [Fact]
    public void ReplaceCrossRunThreeRuns()
    {
        var path = TestHelper.CreateMultiRunDocx(["Hel", "lo Wor", "ld"]);
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "Hello World", "Hi Earth");
        doc.Save();
        doc.Dispose();

        Assert.Equal(1, count);
        var text = TestHelper.ReadDocxText(path);
        Assert.Equal("Hi Earth", text);

        File.Delete(path);
    }

    [Fact]
    public void ReplaceNoMatch()
    {
        var path = TestHelper.CreateTestDocx("Hello World");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "xyz", "abc");
        doc.Dispose();

        Assert.Equal(0, count);
        File.Delete(path);
    }

    [Fact]
    public void ReplaceAllMultiple()
    {
        var path = TestHelper.CreateTestDocx("cat and cat and cat");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "cat", "dog", replaceAll: true);
        doc.Save();
        doc.Dispose();

        Assert.Equal(3, count);
        var text = TestHelper.ReadDocxText(path);
        Assert.Equal("dog and dog and dog", text);

        File.Delete(path);
    }

    [Fact]
    public void ReplaceSingleWhenMultipleExist()
    {
        var path = TestHelper.CreateTestDocx("cat and cat");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.ReplaceInBody(body, "cat", "dog", replaceAll: false);
        doc.Save();
        doc.Dispose();

        Assert.Equal(1, count);

        File.Delete(path);
    }

    [Fact]
    public void CountInBody()
    {
        var path = TestHelper.CreateTestDocx("Hello World", "Hello Again", "Goodbye");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        var count = TextReplacer.CountInBody(body, "Hello");
        doc.Dispose();

        Assert.Equal(2, count);
        File.Delete(path);
    }

    [Fact]
    public void ReplacePreservesFormatting()
    {
        var props = new RunProperties(new Bold());
        var path = TestHelper.CreateMultiRunDocx(["Hello World"], props);

        {
            using var doc = WordprocessingDocument.Open(path, true);
            var body = doc.MainDocumentPart!.Document.Body!;
            TextReplacer.ReplaceInBody(body, "World", "Earth");
            doc.Save();
        }

        // Reopen and check formatting
        using var doc2 = WordprocessingDocument.Open(path, false);
        var body2 = doc2.MainDocumentPart!.Document.Body!;
        var runs = body2.Descendants<Run>().ToList();
        Assert.True(runs.Count > 0);
        // The run with replacement should still have bold
        var hasText = runs.Where(r => r.InnerText.Contains("Earth")).ToList();
        Assert.NotEmpty(hasText);

        File.Delete(path);
    }

    [Fact]
    public void ReplaceWithEmptyStringDeletesText()
    {
        var path = TestHelper.CreateTestDocx("Hello Beautiful World");
        using var doc = WordprocessingDocument.Open(path, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        TextReplacer.ReplaceInBody(body, "Beautiful ", "");
        doc.Save();
        doc.Dispose();

        var text = TestHelper.ReadDocxText(path);
        Assert.Equal("Hello World", text);

        File.Delete(path);
    }
}
