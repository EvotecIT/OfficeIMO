using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Creates a deck of title-and-text slides and saves it, through the Open XML SDK and through the presentation model
/// on the Open XML layer. Both write the same parts with the same markup: one master, one layout, the Office theme,
/// table styles, presentation properties and document properties. Setup checks that both decks list the same parts
/// and slide text.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointModelCreateBenchmarks {
    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() => PowerPointModelWorkload.CheckEquivalent(OpenXmlSdk(), PackageModel());

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() => PowerPointModelWorkload.CreateWithSdk(Slides);

    [Benchmark]
    public byte[] PackageModel() => PowerPointModelWorkload.CreateWithModel(Slides);
}

/// <summary>
/// Opens a saved deck, replaces the title of the middle slide, adds a text box to it and saves the package, through
/// the Open XML SDK and through the presentation model. Setup checks that both results have the same slide text.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointModelEditBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointModelWorkload.CreateWithSdk(Slides);
        PowerPointModelWorkload.CheckEquivalent(OpenXmlSdk(), PackageModel());
    }

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() {
        var stream = new MemoryStream();
        stream.Write(_deck, 0, _deck.Length);
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            P.SlideId entry = document.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>().ElementAt(Slides / 2);
            var slide = (SlidePart)document.PresentationPart.GetPartById(entry.RelationshipId!.Value!);
            P.ShapeTree shapes = slide.Slide!.CommonSlideData!.ShapeTree!;
            P.Shape title = shapes.Elements<P.Shape>().First(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Title);
            A.Paragraph[] paragraphs = title.TextBody!.Elements<A.Paragraph>().ToArray();
            foreach (A.Paragraph paragraph in paragraphs) paragraph.Remove();
            title.TextBody.Append(PowerPointModelWorkload.SdkParagraph("Edited title"));
            uint next = shapes.Elements<P.Shape>().Max(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value) + 1;
            shapes.Append(PowerPointModelWorkload.SdkTextBox(next, 457200, 4572000, 6096000, 914400, "Added later"));
        }
        return stream.ToArray();
    }

    [Benchmark]
    public byte[] PackageModel() {
        using var source = new MemoryStream(_deck, writable: false);
        using PowerPointPackageModel model = PowerPointPackageModel.Open(source);
        PowerPointModelSlide slide = model.GetSlide(Slides / 2);
        PowerPointModelSlide.SetText(slide.FindTitle()!.Value, "Edited title");
        slide.AddTextBox(457200, 4572000, 6096000, 914400, "Added later");
        var output = new MemoryStream();
        model.Save(output);
        return output.ToArray();
    }
}

internal static class PowerPointModelWorkload {
    private const string RelationshipTypes = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private static readonly byte[] Theme = Resource("OfficeIMO.PowerPoint.Resources.theme1.xml");
    private static readonly byte[] TableStyles = Resource("OfficeIMO.PowerPoint.Resources.tableStyles.xml");

    internal static byte[] CreateWithModel(int slides) {
        using PowerPointPackageModel model = PowerPointPackageModel.Create();
        for (int index = 1; index <= slides; index++) {
            PowerPointModelSlide slide = model.AddSlide();
            slide.AddTitle("Slide " + index);
            slide.AddTextBox(838200, 1825625, 10515600, 4351338, "First point of slide " + index, "Second point", "Third & last <point>");
        }
        var output = new MemoryStream();
        model.Save(output);
        return output.ToArray();
    }

    internal static byte[] CreateWithSdk(int slides) {
        var stream = new MemoryStream();
        using (PresentationDocument document = PresentationDocument.Create(stream, PresentationDocumentType.Presentation)) {
            CoreFilePropertiesPart core = document.AddCoreFilePropertiesPart();
            using (Stream output = core.GetStream(FileMode.Create)) {
                byte[] xml = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>OfficeIMO</dc:creator><dcterms:created xsi:type=\"dcterms:W3CDTF\">"
                    + DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", System.Globalization.CultureInfo.InvariantCulture) + "</dcterms:created></cp:coreProperties>");
                output.Write(xml, 0, xml.Length);
            }
            ExtendedFilePropertiesPart app = document.AddExtendedFilePropertiesPart();
            app.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties(new DocumentFormat.OpenXml.ExtendedProperties.Application("OfficeIMO"));

            PresentationPart presentationPart = document.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation(
                new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rId1" }),
                new P.SlideIdList(),
                new P.SlideSize { Cx = 12192000, Cy = 6858000 },
                new P.NotesSize { Cx = 6858000, Cy = 9144000 }) { SaveSubsetFonts = true };

            SlideMasterPart master = presentationPart.AddNewPart<SlideMasterPart>("rId1");
            ThemePart theme = presentationPart.AddNewPart<ThemePart>("rId2");
            using (Stream output = theme.GetStream(FileMode.Create)) output.Write(Theme, 0, Theme.Length);
            PresentationPropertiesPart presProps = presentationPart.AddNewPart<PresentationPropertiesPart>("rId3");
            presProps.PresentationProperties = new P.PresentationProperties();
            TableStylesPart tableStyles = presentationPart.AddNewPart<TableStylesPart>("rId4");
            using (Stream output = tableStyles.GetStream(FileMode.Create)) output.Write(TableStyles, 0, TableStyles.Length);

            master.SlideMaster = new P.SlideMaster(
                new P.CommonSlideData(ShapeTree(
                    Placeholder(2, "Title Placeholder 1", P.PlaceholderValues.Title, null, new[] { 838200L, 365125, 10515600, 1325563 }),
                    Placeholder(3, "Text Placeholder 2", P.PlaceholderValues.Body, 1, new[] { 838200L, 1825625, 10515600, 4351338 }))),
                new P.ColorMap {
                    Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1,
                    Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2,
                    Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2,
                    Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4,
                    Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = "rId1" }));
            SlideLayoutPart layout = master.AddNewPart<SlideLayoutPart>("rId1");
            master.AddPart(theme, "rId2");
            layout.SlideLayout = new P.SlideLayout(
                new P.CommonSlideData(ShapeTree(
                    Placeholder(2, "Title 1", P.PlaceholderValues.Title, null, null),
                    Placeholder(3, "Content Placeholder 2", null, 1, null))) { Name = "Title and Content" },
                new P.ColorMapOverride(new A.MasterColorMapping())) { Type = P.SlideLayoutValues.Object, Preserve = true };
            layout.AddPart(master, "rId1");

            P.SlideIdList slideIds = presentationPart.Presentation.SlideIdList!;
            for (int index = 1; index <= slides; index++) {
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>("rId" + (index + 4));
                slidePart.Slide = new P.Slide(
                    new P.CommonSlideData(ShapeTree(
                        Title(2, "Slide " + index),
                        SdkTextBox(3, 838200, 1825625, 10515600, 4351338, "First point of slide " + index, "Second point", "Third & last <point>"))),
                    new P.ColorMapOverride(new A.MasterColorMapping()));
                slidePart.AddPart(layout, "rId1");
                slideIds.Append(new P.SlideId { Id = (uint)(255 + index), RelationshipId = "rId" + (index + 4) });
            }
        }
        return stream.ToArray();
    }

    internal static P.Shape SdkTextBox(uint id, long x, long y, long cx, long cy, params string[] paragraphs) {
        var body = new P.TextBody(new A.BodyProperties { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false }, new A.ListStyle());
        foreach (string paragraph in paragraphs) body.Append(SdkParagraph(paragraph));
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = "TextBox " + (id - 1) },
                new P.NonVisualShapeDrawingProperties { TextBox = true },
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()),
            body);
    }

    internal static A.Paragraph SdkParagraph(string text) =>
        new(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)));

    private static P.Shape Title(uint id, string text) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = "Title " + (id - 1) },
            new P.NonVisualShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
        new P.ShapeProperties(),
        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), SdkParagraph(text)));

    private static P.Shape Placeholder(uint id, string name, P.PlaceholderValues? type, uint? index, long[]? box) {
        var placeholder = new P.PlaceholderShape();
        if (type.HasValue) placeholder.Type = type.Value;
        if (index.HasValue) placeholder.Index = index.Value;
        var properties = new P.ShapeProperties();
        if (box != null) {
            properties.Append(new A.Transform2D(new A.Offset { X = box[0], Y = box[1] }, new A.Extents { Cx = box[2], Cy = box[3] }));
            properties.Append(new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle });
        }
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(placeholder)),
            properties,
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));
    }

    private static P.ShapeTree ShapeTree(params OpenXmlElement[] shapes) {
        var tree = new P.ShapeTree(GroupProperties());
        tree.Append(shapes);
        return tree;
    }

    private static OpenXmlElement[] GroupProperties() => new OpenXmlElement[] {
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup(
            new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 0, Cy = 0 },
            new A.ChildOffset { X = 0, Y = 0 }, new A.ChildExtents { Cx = 0, Cy = 0 }))
    };

    /// <summary>Checks that two decks have the same parts and the same paragraphs on every slide, in order.</summary>
    internal static void CheckEquivalent(byte[] sdk, byte[] model) {
        List<string> sdkParts = PartNames(sdk);
        List<string> modelParts = PartNames(model);
        if (!sdkParts.SequenceEqual(modelParts)) {
            throw new InvalidOperationException("The decks have different parts: " + string.Join(", ", sdkParts.Except(modelParts)) + " | " + string.Join(", ", modelParts.Except(sdkParts)));
        }
        if (!SlideText(sdk).SequenceEqual(SlideText(model))) throw new InvalidOperationException("The decks have different slide text.");
    }

    private static List<string> PartNames(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        var names = new List<string>();
        var seen = new HashSet<Uri>();
        var pending = new Queue<OpenXmlPart>(document.Parts.Select(pair => pair.OpenXmlPart));
        while (pending.Count > 0) {
            OpenXmlPart part = pending.Dequeue();
            if (!seen.Add(part.Uri)) continue;
            names.Add(part.Uri.OriginalString);
            foreach (IdPartPair child in part.Parts) pending.Enqueue(child.OpenXmlPart);
        }
        names.Sort(StringComparer.Ordinal);
        return names;
    }

    private static List<string> SlideText(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        var text = new List<string>();
        foreach (P.SlideId entry in document.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>()) {
            var slide = (SlidePart)document.PresentationPart.GetPartById(entry.RelationshipId!.Value!);
            text.AddRange(slide.Slide!.Descendants<A.Paragraph>().Select(paragraph => paragraph.InnerText));
        }
        return text;
    }

    private static byte[] Resource(string name) {
        using Stream stream = typeof(OfficeIMO.PowerPoint.PowerPointDocument).Assembly.GetManifestResourceStream(name)
            ?? throw new InvalidOperationException("Missing embedded resource '" + name + "'.");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}
