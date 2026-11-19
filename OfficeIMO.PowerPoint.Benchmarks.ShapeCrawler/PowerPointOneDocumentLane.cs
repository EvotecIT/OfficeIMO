using RecursiveAI.OneDocument;
using RecursiveAI.OneDocument.Presentations;
using DrawingText = RecursiveAI.OneDocument.Schema.Drawing;
using ShapeView = RecursiveAI.OneDocument.Schema.Presentation.Shape;
using OneDocumentPresentation = RecursiveAI.OneDocument.Presentations.Presentation;
using OneDocumentShape = RecursiveAI.OneDocument.Presentations.Shape;
using OneDocumentSlide = RecursiveAI.OneDocument.Presentations.Slide;

namespace OfficeIMO.PowerPoint.Benchmarks.Comparison;

/// <summary>
/// OneDocument implementations of the comparison workloads. OneDocument's slide API adds tables and pictures
/// directly; it has no text box call, so text boxes are built through its typed views over the slide markup,
/// the route its documentation gives for markup the hub does not cover.
/// </summary>
internal static class PowerPointOneDocumentLane {
    private const long Emu = 12700L;
    private const string BlankLayout = "Blank";

    internal static byte[] Create(int slideCount) {
        using OneDocumentPresentation presentation = OneDocumentPresentation.Create();
        SlideLayout blank = presentation.Masters.SelectMany(master => master.Layouts).First(layout => layout.Name == BlankLayout);
        for (int index = 0; index < slideCount; index++) {
            OneDocumentSlide slide = presentation.AddSlide(blank);
            AddTextBox(slide, 2, PowerPointComparisonCorpus.TitleName, 40, 24, 760, 50,
                PowerPointComparisonCorpus.Title(index), bold: true, fontSizeHundredths: 2800);
            ShapeView body = AddTextBox(slide, 3, PowerPointComparisonCorpus.BodyName, 40, 90, 420, 160,
                PowerPointComparisonCorpus.BodyLine(index, 0), bold: false, fontSizeHundredths: null);
            DrawingText.TextBody bodyText = body.TxBody!.Value;
            for (int line = 1; line < PowerPointComparisonCorpus.BodyLineCount; line++) {
                AddParagraph(bodyText, PowerPointComparisonCorpus.BodyLine(index, line), bold: false, fontSizeHundredths: null);
            }

            var rows = new string[PowerPointComparisonCorpus.TableRows][];
            for (int row = 0; row < rows.Length; row++) {
                rows[row] = new string[PowerPointComparisonCorpus.TableColumns];
                for (int column = 0; column < rows[row].Length; column++) {
                    rows[row][column] = PowerPointComparisonCorpus.Cell(index, row, column);
                }
            }
            slide.AddTable(rows, 500, 90, 360, 120.0 / PowerPointComparisonCorpus.TableRows);
            slide.AddPicture(PowerPointComparisonCorpus.Image, 40, 300, 230, 54, altText: null);
        }
        return Save(presentation);
    }

    internal static PowerPointTextObservation Read(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentPresentation presentation = OneDocumentPresentation.Open(stream);
        PowerPointTextObservation observation = PowerPointTextObservation.Empty;
        foreach (OneDocumentSlide slide in presentation.Slides) {
            foreach (OneDocumentShape shape in slide.Shapes) {
                if (shape.Table is { } table) {
                    foreach (TableRow row in table.Rows) {
                        foreach (TableCell cell in row.Cells) observation = observation.Add(cell.Text);
                    }
                } else if (shape.TextBody is { } textBody) {
                    observation = observation.Add(textBody.Text);
                }
            }
        }
        return observation;
    }

    internal static byte[] Edit(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentPresentation presentation = OneDocumentPresentation.Open(stream);
        int index = 0;
        foreach (OneDocumentSlide slide in presentation.Slides) {
            OneDocumentShape title = slide.Shapes.First(shape => shape.Name == PowerPointComparisonCorpus.TitleName);
            title.TextBody!.SetText(PowerPointComparisonCorpus.EditedTitle(index++));
            slide.Shapes.First(shape => shape.Table != null).Table![1, 1].SetText(PowerPointComparisonCorpus.EditedCell);
        }
        return Save(presentation);
    }

    internal static byte[] CopySlides(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using OneDocumentPresentation presentation = OneDocumentPresentation.Open(stream);
        OneDocumentSlide[] originals = presentation.Slides.ToArray();
        foreach (OneDocumentSlide slide in originals) presentation.DuplicateSlide(slide, presentation.Slides.Count);
        return Save(presentation);
    }

    private static ShapeView AddTextBox(OneDocumentSlide slide, long id, string name, long left, long top, long width, long height,
        string text, bool bold, long? fontSizeHundredths) {
        ShapeView shape = slide.View!.Value.CSld!.Value.SpTree!.Value.AddSp();
        var nonVisual = shape.GetOrAddNvSpPr();
        DrawingText.NonVisualDrawingProps properties = nonVisual.GetOrAddCNvPr();
        properties.Id = id;
        properties.Name = name;
        nonVisual.GetOrAddCNvSpPr().TxBox = true;
        nonVisual.GetOrAddNvPr();

        DrawingText.ShapeProperties shapeProperties = shape.GetOrAddSpPr();
        DrawingText.Transform2D transform = shapeProperties.GetOrAddTransform();
        DrawingText.Point2D offset = transform.GetOrAddOff();
        offset.X = left * Emu;
        offset.Y = top * Emu;
        DrawingText.PositiveSize2D extent = transform.GetOrAddExtension();
        extent.Cx = width * Emu;
        extent.Cy = height * Emu;
        shapeProperties.GetOrAddPresetGeometry().Prst = DrawingText.ShapeTypeValue.Rect;

        DrawingText.TextBody textBody = shape.GetOrAddTxBody();
        textBody.GetOrAddBodyProperties().Wrap = DrawingText.TextWrappingTypeValue.Square;
        textBody.GetOrAddListStyle();
        AddParagraph(textBody, text, bold, fontSizeHundredths);
        return shape;
    }

    private static void AddParagraph(DrawingText.TextBody textBody, string text, bool bold, long? fontSizeHundredths) {
        DrawingText.RegularTextRun run = textBody.AddParagraph().AddRun();
        DrawingText.TextCharacterProperties runProperties = run.GetOrAddRunProperties();
        runProperties.Lang = "en-US";
        if (bold) runProperties.B = true;
        if (fontSizeHundredths is { } size) runProperties.Sz = size;
        run.AddT().Node.SetText(text);
    }

    private static byte[] Save(OneDocumentPresentation presentation) {
        using var output = new MemoryStream();
        presentation.SaveAs(output, OfficeFormat.Pptx);
        return output.ToArray();
    }
}
