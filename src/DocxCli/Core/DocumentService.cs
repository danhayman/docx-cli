using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxCli.Core;

public static class DocumentService
{
    public static WordprocessingDocument OpenRead(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        return WordprocessingDocument.Open(path, false);
    }

    public static WordprocessingDocument OpenForEdit(string path, string? outputPath, bool backup)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        if (backup)
        {
            var bakPath = path + ".bak";
            File.Copy(path, bakPath, overwrite: true);
            Console.Error.WriteLine($"backed up to {bakPath}");
        }

        // If writing to a different output, clone there and edit the clone
        if (outputPath != null)
        {
            File.Copy(path, outputPath, overwrite: true);
            return WordprocessingDocument.Open(outputPath, true);
        }

        // In-place edit: clone to temp, we'll move back after save
        return WordprocessingDocument.Open(path, true);
    }

    public static void SaveAtomically(WordprocessingDocument doc, string originalPath, string? outputPath)
    {
        // If outputPath was used, the doc is already editing the output file — just save
        if (outputPath != null)
        {
            doc.Save();
            doc.Dispose();
            return;
        }

        // For in-place edits, we're editing the original directly
        doc.Save();
        doc.Dispose();
    }

    public static List<ParagraphInfo> GetParagraphs(Body body)
    {
        var result = new List<ParagraphInfo>();
        int index = 1;

        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph para)
            {
                result.Add(new ParagraphInfo(index++, para));
            }
            else if (element is Table table)
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    var cellTexts = new List<string>();
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        var text = string.Join(" ", cell.Elements<Paragraph>().Select(p => p.InnerText));
                        cellTexts.Add(text);
                    }
                    var tableParaText = string.Join("\t", cellTexts);
                    // Create a synthetic paragraph info for table rows
                    result.Add(new ParagraphInfo(index++, null, tableParaText, IsTableRow: true));
                }
            }
        }

        return result;
    }

    public static int CountWords(Body body)
    {
        int count = 0;
        foreach (var para in body.Descendants<Paragraph>())
        {
            var text = para.InnerText;
            if (!string.IsNullOrWhiteSpace(text))
                count += text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        return count;
    }
}

public record ParagraphInfo(int Number, Paragraph? Paragraph, string? OverrideText = null, bool IsTableRow = false)
{
    public string Text => OverrideText ?? Paragraph?.InnerText ?? "";
}
