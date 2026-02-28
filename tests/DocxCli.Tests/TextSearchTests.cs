using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxCli.Core;

namespace DocxCli.Tests;

public class TextSearchTests
{
    [Fact]
    public void FindInSingleRun()
    {
        var para = new Paragraph(new Run(new Text("Hello World")));
        var matches = TextSearch.FindInParagraph(para, "World");

        Assert.Single(matches);
        Assert.Equal(6, matches[0].StartInConcat);
        Assert.Equal(5, matches[0].Length);
        Assert.Equal(0, matches[0].StartRunIndex);
        Assert.Equal(6, matches[0].StartOffsetInRun);
        Assert.Equal(0, matches[0].EndRunIndex);
        Assert.Equal(11, matches[0].EndOffsetInRun);
    }

    [Fact]
    public void FindAcrossTwoRuns()
    {
        var para = new Paragraph(
            new Run(new Text("Hello Wo")),
            new Run(new Text("rld!")));

        var matches = TextSearch.FindInParagraph(para, "World");
        Assert.Single(matches);
        Assert.Equal(0, matches[0].StartRunIndex);
        Assert.Equal(6, matches[0].StartOffsetInRun);
        Assert.Equal(1, matches[0].EndRunIndex);
        Assert.Equal(3, matches[0].EndOffsetInRun);
    }

    [Fact]
    public void FindAcrossThreeRuns()
    {
        var para = new Paragraph(
            new Run(new Text("Hel")),
            new Run(new Text("lo Wor")),
            new Run(new Text("ld")));

        var matches = TextSearch.FindInParagraph(para, "Hello World");
        Assert.Single(matches);
        Assert.Equal(0, matches[0].StartRunIndex);
        Assert.Equal(0, matches[0].StartOffsetInRun);
        Assert.Equal(2, matches[0].EndRunIndex);
        Assert.Equal(2, matches[0].EndOffsetInRun);
    }

    [Fact]
    public void FindMultipleMatches()
    {
        var para = new Paragraph(new Run(new Text("cat and cat and cat")));
        var matches = TextSearch.FindInParagraph(para, "cat");
        Assert.Equal(3, matches.Count);
    }

    [Fact]
    public void FindNoMatch()
    {
        var para = new Paragraph(new Run(new Text("Hello World")));
        var matches = TextSearch.FindInParagraph(para, "xyz");
        Assert.Empty(matches);
    }

    [Fact]
    public void FindCaseInsensitive()
    {
        var para = new Paragraph(new Run(new Text("Hello World")));
        var matches = TextSearch.FindInParagraph(para, "hello", StringComparison.OrdinalIgnoreCase);
        Assert.Single(matches);
    }

    [Fact]
    public void FindEmptyParagraph()
    {
        var para = new Paragraph();
        var matches = TextSearch.FindInParagraph(para, "test");
        Assert.Empty(matches);
    }

    [Fact]
    public void FindEmptySearch()
    {
        var para = new Paragraph(new Run(new Text("Hello")));
        var matches = TextSearch.FindInParagraph(para, "");
        Assert.Empty(matches);
    }

    [Fact]
    public void FindAtRunBoundary()
    {
        // Match starts exactly at the beginning of a run
        var para = new Paragraph(
            new Run(new Text("Hello ")),
            new Run(new Text("World")));

        var matches = TextSearch.FindInParagraph(para, "World");
        Assert.Single(matches);
        Assert.Equal(1, matches[0].StartRunIndex);
        Assert.Equal(0, matches[0].StartOffsetInRun);
    }

    [Fact]
    public void FindMatchEndsAtRunBoundary()
    {
        var para = new Paragraph(
            new Run(new Text("Hello")),
            new Run(new Text(" World")));

        var matches = TextSearch.FindInParagraph(para, "Hello");
        Assert.Single(matches);
        Assert.Equal(0, matches[0].StartRunIndex);
        // End offset 5 maps to run 1 offset 0 (start of second run)
        Assert.Equal(1, matches[0].EndRunIndex);
        Assert.Equal(0, matches[0].EndOffsetInRun);
    }
}
