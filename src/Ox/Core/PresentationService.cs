using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace Ox.Core;

public static class PresentationService
{
    public static List<ParagraphInfo> GetParagraphs(PresentationDocument doc)
    {
        var result = new List<ParagraphInfo>();
        int index = 1;

        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("invalid .pptx: no presentation part");

        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>()
            ?? Enumerable.Empty<SlideId>();

        foreach (var slideId in slideIds)
        {
            var relId = slideId.RelationshipId?.Value
                ?? throw new InvalidOperationException("slide missing relationship ID");

            var slidePart = (SlidePart)presentationPart.GetPartById(relId);

            foreach (var shape in slidePart.Slide.Descendants<Shape>())
            {
                var textBody = shape.TextBody;
                if (textBody == null) continue;

                foreach (var para in textBody.Elements<D.Paragraph>())
                {
                    var text = string.Concat(para.Descendants<D.Run>()
                        .Select(r => r.Text?.Text ?? ""));

                    if (!string.IsNullOrEmpty(text))
                        result.Add(new ParagraphInfo(index++, null, text));
                }
            }
        }

        return result;
    }

    public static int CountSlides(PresentationDocument doc)
    {
        return doc.PresentationPart?.Presentation.SlideIdList?.Elements<SlideId>().Count() ?? 0;
    }

    public static int CountWords(PresentationDocument doc)
    {
        int count = 0;
        foreach (var para in GetParagraphs(doc))
        {
            if (!string.IsNullOrWhiteSpace(para.Text))
                count += para.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        }
        return count;
    }
}
