using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ox.Core;

namespace Ox.Tests;

public class IntegrationTests
{
    [Fact]
    public void CatReadsAllParagraphs()
    {
        var path = TestHelper.CreateTestDocx("First paragraph", "", "Third paragraph");
        using var doc = DocumentService.OpenRead(path);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = DocumentService.GetParagraphs(body);

        Assert.Equal(3, paragraphs.Count);
        Assert.Equal("First paragraph", paragraphs[0].Text);
        Assert.Equal("", paragraphs[1].Text);
        Assert.Equal("Third paragraph", paragraphs[2].Text);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void InfoCountsWords()
    {
        var path = TestHelper.CreateTestDocx("Hello World", "One two three");
        using var doc = DocumentService.OpenRead(path);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var words = DocumentService.CountWords(body);
        Assert.Equal(5, words);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void TableExtraction()
    {
        var path = TestHelper.CreateDocxWithTable(
            ["Header"],
            [["A", "B"], ["C", "D"]]);

        using var doc = DocumentService.OpenRead(path);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = DocumentService.GetParagraphs(body);

        Assert.Equal(3, paragraphs.Count); // 1 para + 2 table rows
        Assert.Equal("Header", paragraphs[0].Text);
        Assert.Contains("A", paragraphs[1].Text);
        Assert.Contains("B", paragraphs[1].Text);
        Assert.True(paragraphs[1].IsTableRow);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void LockFileDetection()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"lock_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var docPath = Path.Combine(dir, "test.docx");
        var lockPath = Path.Combine(dir, "~$test.docx");

        File.WriteAllText(docPath, "dummy");
        Assert.Null(LockDetector.GetLockFile(docPath));

        File.WriteAllText(lockPath, "lock");
        Assert.NotNull(LockDetector.GetLockFile(docPath));

        // Force should not throw
        LockDetector.CheckLock(docPath, force: true);

        // Without force should throw
        Assert.Throws<InvalidOperationException>(() => LockDetector.CheckLock(docPath, force: false));

        Directory.Delete(dir, true);
    }

    [Fact]
    public void RoundTripEditAndRead()
    {
        var path = TestHelper.CreateTestDocx("The term is 30 days from signing.");
        var outputPath = path + ".edited.docx";

        // Edit
        using (var doc = DocumentService.OpenForEdit(path, outputPath, backup: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            TextReplacer.ReplaceInBody(body, "30 days", "60 days");
            DocumentService.SaveAtomically(doc, path, outputPath);
        }

        // Read back
        var paragraphs = TestHelper.ReadDocxParagraphs(outputPath);
        Assert.Single(paragraphs);
        Assert.Equal("The term is 60 days from signing.", paragraphs[0]);

        File.Delete(path);
        File.Delete(outputPath);
    }

    [Fact]
    public void BackupCreation()
    {
        var path = TestHelper.CreateTestDocx("Original text");
        var bakPath = path + ".bak";

        using var doc = DocumentService.OpenForEdit(path, null, backup: true);
        DocumentService.SaveAtomically(doc, path, null);

        Assert.True(File.Exists(bakPath));
        File.Delete(path);
        File.Delete(bakPath);
    }

    [Fact]
    public void EmptyDocument()
    {
        var path = TestHelper.CreateTestDocx();
        using var doc = DocumentService.OpenRead(path);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = DocumentService.GetParagraphs(body);
        var words = DocumentService.CountWords(body);

        Assert.Empty(paragraphs);
        Assert.Equal(0, words);

        doc.Dispose();
        File.Delete(path);
    }

    [Fact]
    public void ZeroByteFile_ThrowsMeaningfulError()
    {
        var path = Path.GetTempFileName();
        // File is already zero bytes

        var ex = Assert.Throws<InvalidOperationException>(() => DocumentService.OpenRead(path));
        Assert.Contains("empty", ex.Message);

        File.Delete(path);
    }

    [Fact]
    public void OleCompoundDocument_ThrowsMeaningfulError()
    {
        var path = Path.GetTempFileName();
        // Write OLE compound document magic bytes
        File.WriteAllBytes(path, [0xD0, 0xCF, 0x11, 0xE0, 0x00, 0x00, 0x00, 0x00]);

        var ex = Assert.Throws<InvalidOperationException>(() => DocumentService.OpenRead(path));
        Assert.Contains("password-protected", ex.Message);

        File.Delete(path);
    }
}
