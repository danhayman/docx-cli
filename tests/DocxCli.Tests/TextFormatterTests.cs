using DocxCli.Output;

namespace DocxCli.Tests;

public class TextFormatterTests
{
    [Fact]
    public void WrapTextShortLine()
    {
        var lines = TextFormatter.WrapText("Hello World", 80);
        Assert.Single(lines);
        Assert.Equal("Hello World", lines[0]);
    }

    [Fact]
    public void WrapTextLongLine()
    {
        var lines = TextFormatter.WrapText("one two three four five six", 15);
        Assert.True(lines.Count >= 2);
        foreach (var line in lines)
        {
            Assert.True(line.Length <= 15 || !line.Contains(' '));
        }
    }

    [Fact]
    public void WrapTextEmptyString()
    {
        var lines = TextFormatter.WrapText("", 80);
        Assert.Single(lines);
        Assert.Equal("", lines[0]);
    }

    [Fact]
    public void WrapTextSingleLongWord()
    {
        var lines = TextFormatter.WrapText("superlongword", 5);
        Assert.Single(lines);
        Assert.Equal("superlongword", lines[0]); // Can't break a word
    }

    [Fact]
    public void WrapTextExactWidth()
    {
        var lines = TextFormatter.WrapText("Hello World", 11);
        Assert.Single(lines);
        Assert.Equal("Hello World", lines[0]);
    }
}
