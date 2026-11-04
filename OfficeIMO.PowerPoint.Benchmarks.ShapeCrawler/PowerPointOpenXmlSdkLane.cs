using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks.Comparison;

/// <summary>Format-native Open XML SDK implementations of the comparison workloads.</summary>
internal static class PowerPointOpenXmlSdkLane {
    private const long Emu = 12700L;

    internal static byte[] Create(int slideCount) {
        using MemoryStream stream = PowerPointComparisonCorpus.Editable(PowerPointComparisonCorpus.EmptyShell);
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            PresentationPart presentationPart = document.PresentationPart!;
            P.Presentation presentation = presentationPart.Presentation!;
            SlideLayoutPart layout = presentationPart.SlideMasterParts.First().SlideLayoutParts.First();
            P.SlideIdList slideIds = presentation.SlideIdList ??= new P.SlideIdList();
            uint nextId = 256;
            for (int index = 0; index < slideCount; index++) {
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.AddPart(layout);
                ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
                using (var image = new MemoryStream(PowerPointComparisonCorpus.Image, writable: false)) {
                    imagePart.FeedData(image);
                }
                slidePart.Slide = CreateSlide(index, slidePart.GetIdOfPart(imagePart));
                slideIds.Append(new P.SlideId { Id = nextId++, RelationshipId = presentationPart.GetIdOfPart(slidePart) });
            }
            presentation.Save();
        }
        return stream.ToArray();
    }

    internal static PowerPointTextObservation Read(byte[] fixture) {
        using var stream = new MemoryStream(fixture, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        PresentationPart presentationPart = document.PresentationPart!;
        PowerPointTextObservation observation = PowerPointTextObservation.Empty;
        foreach (P.SlideId id in presentationPart.Presentation!.SlideIdList!.Elements<P.SlideId>()) {
            var slidePart = (SlidePart)presentationPart.GetPartById(id.RelationshipId!.Value!);
            P.ShapeTree tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
            foreach (P.Shape shape in tree.Elements<P.Shape>()) {
                if (shape.TextBody != null) observation = observation.Add(shape.TextBody.InnerText);
            }
            foreach (A.TableCell cell in tree.Descendants<A.TableCell>()) {
                observation = observation.Add(cell.InnerText);
            }
        }
        return observation;
    }

    internal static byte[] Edit(byte[] fixture) {
        using MemoryStream stream = PowerPointComparisonCorpus.Editable(fixture);
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            PresentationPart presentationPart = document.PresentationPart!;
            int index = 0;
            foreach (P.SlideId id in presentationPart.Presentation!.SlideIdList!.Elements<P.SlideId>()) {
                var slidePart = (SlidePart)presentationPart.GetPartById(id.RelationshipId!.Value!);
                P.ShapeTree tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
                P.Shape title = tree.Elements<P.Shape>().First(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == PowerPointComparisonCorpus.TitleName);
                title.TextBody!.Descendants<A.Text>().First().Text = PowerPointComparisonCorpus.EditedTitle(index);
                A.TableCell cell = tree.Descendants<A.TableRow>().ElementAt(1).Elements<A.TableCell>().ElementAt(1);
                cell.Descendants<A.Text>().First().Text = PowerPointComparisonCorpus.EditedCell;
                slidePart.Slide.Save();
                index++;
            }
        }
        return stream.ToArray();
    }

    internal static byte[] CopySlides(byte[] fixture) {
        using MemoryStream stream = PowerPointComparisonCorpus.Editable(fixture);
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            PresentationPart presentationPart = document.PresentationPart!;
            P.SlideIdList slideIds = presentationPart.Presentation!.SlideIdList!;
            P.SlideId[] originals = slideIds.Elements<P.SlideId>().ToArray();
            uint nextId = originals.Max(id => id.Id!.Value) + 1;
            foreach (P.SlideId original in originals) {
                var source = (SlidePart)presentationPart.GetPartById(original.RelationshipId!.Value!);
                SlidePart copy = presentationPart.AddNewPart<SlidePart>();
                using (Stream sourceXml = source.GetStream(FileMode.Open, FileAccess.Read)) {
                    copy.FeedData(sourceXml);
                }
                foreach (IdPartPair pair in source.Parts) {
                    copy.AddPart(pair.OpenXmlPart, pair.RelationshipId);
                }
                slideIds.Append(new P.SlideId { Id = nextId++, RelationshipId = presentationPart.GetIdOfPart(copy) });
            }
            presentationPart.Presentation.Save();
        }
        return stream.ToArray();
    }

    private static P.Slide CreateSlide(int index, string imageRelationshipId) {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));
        tree.Append(TextShape(2U, PowerPointComparisonCorpus.TitleName, 40, 24, 760, 50,
            Paragraph(PowerPointComparisonCorpus.Title(index), 2800, bold: true)));
        tree.Append(TextShape(3U, PowerPointComparisonCorpus.BodyName, 40, 90, 420, 160,
            Enumerable.Range(0, PowerPointComparisonCorpus.BodyLineCount)
                .Select(line => Paragraph(PowerPointComparisonCorpus.BodyLine(index, line), 1600, bold: false))
                .ToArray()));
        tree.Append(Table(4U, index));
        tree.Append(Picture(5U, imageRelationshipId));
        return new P.Slide(
            new P.CommonSlideData(tree),
            new P.ColorMapOverride(new A.MasterColorMapping()));
    }

    private static P.Shape TextShape(uint id, string name, long x, long y, long width, long height,
        params A.Paragraph[] paragraphs) {
        var body = new P.TextBody(new A.BodyProperties { Wrap = A.TextWrappingValues.Square }, new A.ListStyle());
        body.Append(paragraphs);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties { TextBox = true },
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                Transform(x, y, width, height),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()),
            body);
    }

    private static A.Paragraph Paragraph(string text, int size, bool bold) =>
        new(new A.Run(
            new A.RunProperties { Language = "en-US", FontSize = size, Bold = bold, Dirty = false },
            new A.Text(text)));

    private static P.GraphicFrame Table(uint id, int slide) {
        const long columnWidth = 120L * Emu;
        var grid = new A.TableGrid();
        for (int column = 0; column < PowerPointComparisonCorpus.TableColumns; column++) {
            grid.Append(new A.GridColumn { Width = columnWidth });
        }
        var table = new A.Table(new A.TableProperties { FirstRow = true, BandRow = true }, grid);
        for (int row = 0; row < PowerPointComparisonCorpus.TableRows; row++) {
            var tableRow = new A.TableRow { Height = 30L * Emu };
            for (int column = 0; column < PowerPointComparisonCorpus.TableColumns; column++) {
                tableRow.Append(new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(
                            new A.RunProperties { Language = "en-US", Dirty = false },
                            new A.Text(PowerPointComparisonCorpus.Cell(slide, row, column))))),
                    new A.TableCellProperties()));
            }
            table.Append(tableRow);
        }
        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = PowerPointComparisonCorpus.TableName },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(new A.Offset { X = 500L * Emu, Y = 90L * Emu }, new A.Extents { Cx = 360L * Emu, Cy = 120L * Emu }),
            new A.Graphic(new A.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
    }

    private static P.Picture Picture(uint id, string relationshipId) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = PowerPointComparisonCorpus.PictureName },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                Transform(40, 300, 230, 54),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

    private static A.Transform2D Transform(long x, long y, long width, long height) =>
        new(new A.Offset { X = x * Emu, Y = y * Emu }, new A.Extents { Cx = width * Emu, Cy = height * Emu });
}
