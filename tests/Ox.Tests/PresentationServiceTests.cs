using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace Ox.Tests;

public class PresentationServiceTests
{
    [Fact]
    public void GetParagraphs_SingleSlide_ExtractsText()
    {
        var path = CreateTestPptx("Hello World");
        try
        {
            using var doc = PresentationDocument.Open(path, false);
            var paragraphs = Ox.Core.PresentationService.GetParagraphs(doc);

            Assert.Single(paragraphs);
            Assert.Equal("Hello World", paragraphs[0].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_MultipleSlides_ExtractsAll()
    {
        var path = CreateTestPptx("Slide One", "Slide Two", "Slide Three");
        try
        {
            using var doc = PresentationDocument.Open(path, false);
            var paragraphs = Ox.Core.PresentationService.GetParagraphs(doc);

            Assert.Equal(3, paragraphs.Count);
            Assert.Equal("Slide One", paragraphs[0].Text);
            Assert.Equal("Slide Two", paragraphs[1].Text);
            Assert.Equal("Slide Three", paragraphs[2].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_EmptyPresentation_ReturnsEmpty()
    {
        var path = CreateEmptyPptx();
        try
        {
            using var doc = PresentationDocument.Open(path, false);
            var paragraphs = Ox.Core.PresentationService.GetParagraphs(doc);

            Assert.Empty(paragraphs);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GetParagraphs_MultipleShapesPerSlide_ExtractsAll()
    {
        var path = CreatePptxWithMultipleShapes("Title Text", "Body Text");
        try
        {
            using var doc = PresentationDocument.Open(path, false);
            var paragraphs = Ox.Core.PresentationService.GetParagraphs(doc);

            Assert.Equal(2, paragraphs.Count);
            Assert.Equal("Title Text", paragraphs[0].Text);
            Assert.Equal("Body Text", paragraphs[1].Text);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CountSlides_ReturnsCorrectCount()
    {
        var path = CreateTestPptx("One", "Two");
        try
        {
            using var doc = PresentationDocument.Open(path, false);
            Assert.Equal(2, Ox.Core.PresentationService.CountSlides(doc));
        }
        finally { File.Delete(path); }
    }

    // -- Helpers --

    private static string CreateTestPptx(params string[] slideTexts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx_test_{Guid.NewGuid():N}.pptx");
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slideIdList = new SlideIdList();
        presentationPart.Presentation.SlideIdList = slideIdList;

        uint slideId = 256;

        foreach (var text in slideTexts)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(),
                        new Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties { Id = 2, Name = "TextBox" },
                                new NonVisualShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new ShapeProperties(),
                            new TextBody(
                                new D.BodyProperties(),
                                new D.Paragraph(
                                    new D.Run(
                                        new D.RunProperties { Language = "en-US" },
                                        new D.Text(text))))))));

            slideIdList.Append(new SlideId
            {
                Id = slideId++,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        presentationPart.Presentation.Save();
        return path;
    }

    private static string CreateEmptyPptx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx_test_{Guid.NewGuid():N}.pptx");
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        presentationPart.Presentation.SlideIdList = new SlideIdList();
        presentationPart.Presentation.Save();

        return path;
    }

    private static string CreatePptxWithMultipleShapes(params string[] shapeTexts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx_test_{Guid.NewGuid():N}.pptx");
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slideIdList = new SlideIdList();
        presentationPart.Presentation.SlideIdList = slideIdList;

        var slidePart = presentationPart.AddNewPart<SlidePart>();

        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties());

        uint shapeId = 2;
        foreach (var text in shapeTexts)
        {
            shapeTree.Append(new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = shapeId++, Name = $"TextBox{shapeId}" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(),
                new TextBody(
                    new D.BodyProperties(),
                    new D.Paragraph(
                        new D.Run(
                            new D.RunProperties { Language = "en-US" },
                            new D.Text(text))))));
        }

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));

        slideIdList.Append(new SlideId
        {
            Id = 256,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });

        presentationPart.Presentation.Save();
        return path;
    }
}
