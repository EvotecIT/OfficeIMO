using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks.Comparison;

/// <summary>
/// Deterministic deck content shared by every library lane, plus the Open XML
/// package-level oracle that each lane's output must satisfy.
/// </summary>
internal static class PowerPointComparisonCorpus {
    internal const string TitleName = "Title";
    internal const string BodyName = "Body";
    internal const string TableName = "Metrics";
    internal const string PictureName = "Logo";
    internal const int BodyLineCount = 3;
    internal const int TableRows = 4;
    internal const int TableColumns = 3;
    internal const long SlideWidthEmu = 12192000L;
    internal const long SlideHeightEmu = 6858000L;

    private static byte[]? _image;
    private static byte[]? _emptyShell;

    internal static byte[] Image => _image ??= File.ReadAllBytes(
        Path.Combine(AppContext.BaseDirectory, "benchmark-logo.png"));

    /// <summary>A zero-slide package with master, layouts, and theme, used as the Open XML SDK starting point.</summary>
    internal static byte[] EmptyShell => _emptyShell ??= CreateEmptyShell();

    internal static string Title(int slide) =>
        "Slide " + (slide + 1).ToString("D4", CultureInfo.InvariantCulture) + " title";

    internal static string EditedTitle(int slide) =>
        "Reviewed slide " + (slide + 1).ToString("D4", CultureInfo.InvariantCulture);

    internal static string BodyLine(int slide, int line) =>
        "Point " + (line + 1).ToString(CultureInfo.InvariantCulture) + " for slide " +
        (slide + 1).ToString(CultureInfo.InvariantCulture) + " explains the measured result.";

    internal static string Cell(int slide, int row, int column) {
        if (row == 0) return column switch { 0 => "Metric", 1 => "Current", _ => "Target" };
        if (column == 0) return row switch { 1 => "Quality", 2 => "Coverage", _ => "Latency" };
        return ((slide * 7 + row * 3 + column) % 97).ToString(CultureInfo.InvariantCulture);
    }

    internal const string EditedCell = "Updated";

    private static byte[] CreateEmptyShell() {
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            presentation.Save();
        }
        byte[] bytes = stream.ToArray();
        using var check = new MemoryStream(bytes, writable: false);
        using PresentationDocument document = PresentationDocument.Open(check, false);
        int slides = document.PresentationPart?.Presentation?.SlideIdList?.Elements<P.SlideId>().Count() ?? 0;
        if (slides != 0) throw new InvalidDataException("The empty shell unexpectedly contains slides.");
        if (document.PresentationPart!.SlideMasterParts.FirstOrDefault()?.SlideLayoutParts.Any() != true) {
            throw new InvalidDataException("The empty shell has no slide layouts.");
        }
        return bytes;
    }

    /// <summary>Copies bytes into an expandable stream so Open XML SDK can edit in place.</summary>
    internal static MemoryStream Editable(byte[] payload) {
        var stream = new MemoryStream(payload.Length * 2);
        stream.Write(payload, 0, payload.Length);
        stream.Position = 0;
        return stream;
    }

    internal static PresentationExpectation Created(int slideCount) =>
        new(slideCount, Title, slide => Cell(slide, 1, 1), slideCount, RequireNames: false);

    internal static PresentationExpectation Edited(int slideCount) =>
        new(slideCount, EditedTitle, _ => EditedCell, slideCount, RequireNames: true);

    internal static PresentationExpectation Copied(int slideCount) =>
        new(slideCount * 2, slide => Title(slide % slideCount), slide => Cell(slide % slideCount, 1, 1),
            slideCount, RequireNames: true);

    /// <summary>Validates content, picture payloads, and Open XML conformance after the measured operation.</summary>
    internal static void Validate(string lane, byte[] payload, PresentationExpectation expected) {
        using var input = new MemoryStream(payload, writable: false);
        using PresentationDocument document = PresentationDocument.Open(input, false);
        PresentationPart presentationPart = document.PresentationPart
            ?? throw new InvalidDataException(lane + " produced no presentation part.");
        P.SlideId[] ids = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
        Ensure(lane, "slide count", expected.SlideCount, ids.Length);
        for (int index = 0; index < ids.Length; index++) {
            var slidePart = (SlidePart)presentationPart.GetPartById(ids[index].RelationshipId!.Value!);
            ValidateSlide(lane, slidePart, index, expected);
        }
        ValidationErrorInfo? error = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).FirstOrDefault();
        if (error != null) {
            throw new InvalidDataException(lane + " produced invalid Open XML: " + error.Description + " at " + error.Path?.XPath);
        }
    }

    private static void ValidateSlide(string lane, SlidePart slidePart, int index, PresentationExpectation expected) {
        P.ShapeTree tree = slidePart.Slide?.CommonSlideData?.ShapeTree
            ?? throw new InvalidDataException(lane + " slide " + index + " has no shape tree.");
        string prefix = lane + " slide " + (index + 1) + " ";
        int contentSlide = index % expected.ContentSlideCount;

        P.Shape[] shapes = tree.Elements<P.Shape>().ToArray();
        P.Shape title = shapes.SingleOrDefault(shape => ShapeText(shape) == expected.Title(index))
            ?? throw new InvalidDataException(prefix + "has no title '" + expected.Title(index) + "'.");
        if (title.Descendants<A.RunProperties>().FirstOrDefault()?.Bold?.Value != true) {
            throw new InvalidDataException(prefix + "title is not bold.");
        }
        string[] bodyLines = Enumerable.Range(0, BodyLineCount).Select(line => BodyLine(contentSlide, line)).ToArray();
        P.Shape body = shapes.SingleOrDefault(shape =>
                shape.TextBody?.Elements<A.Paragraph>().Select(ParagraphText).SequenceEqual(bodyLines) == true)
            ?? throw new InvalidDataException(prefix + "has no three-paragraph body.");
        if (expected.RequireNames) {
            Ensure(lane, "title name", TitleName, title.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty);
            Ensure(lane, "body name", BodyName, body.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty);
        }

        A.Table table = tree.Descendants<A.Table>().SingleOrDefault()
            ?? throw new InvalidDataException(prefix + "has no single table.");
        A.TableRow[] rows = table.Elements<A.TableRow>().ToArray();
        Ensure(lane, "table rows", TableRows, rows.Length);
        for (int row = 0; row < TableRows; row++) {
            A.TableCell[] cells = rows[row].Elements<A.TableCell>().ToArray();
            Ensure(lane, "table columns", TableColumns, cells.Length);
            for (int column = 0; column < TableColumns; column++) {
                string expectedText = row == 1 && column == 1
                    ? expected.EditedCell(index)
                    : Cell(contentSlide, row, column);
                Ensure(lane, "cell " + row + "," + column + " on slide " + (index + 1), expectedText,
                    string.Concat(cells[column].Descendants<A.Text>().Select(text => text.Text)));
            }
        }

        P.Picture picture = tree.Elements<P.Picture>().SingleOrDefault()
            ?? throw new InvalidDataException(prefix + "has no single picture.");
        string embed = picture.BlipFill?.Blip?.Embed?.Value
            ?? throw new InvalidDataException(prefix + "picture has no embedded blip.");
        var imagePart = (ImagePart)slidePart.GetPartById(embed);
        using Stream imageStream = imagePart.GetStream();
        using var copy = new MemoryStream();
        imageStream.CopyTo(copy);
        if (!copy.ToArray().AsSpan().SequenceEqual(Image)) {
            throw new InvalidDataException(prefix + "picture bytes differ from the source image.");
        }
    }

    private static string ShapeText(P.Shape shape) =>
        string.Join("\n", shape.TextBody?.Elements<A.Paragraph>().Select(ParagraphText) ?? []);

    private static string ParagraphText(A.Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text));

    private static void Ensure(string lane, string name, int expected, int actual) {
        if (expected != actual) throw new InvalidDataException(lane + " " + name + " was " + actual + "; expected " + expected + ".");
    }

    private static void Ensure(string lane, string name, string expected, string actual) {
        if (!string.Equals(expected, actual, StringComparison.Ordinal)) {
            throw new InvalidDataException(lane + " " + name + " was '" + actual + "'; expected '" + expected + "'.");
        }
    }
}

internal sealed record PresentationExpectation(
    int SlideCount,
    Func<int, string> Title,
    Func<int, string> EditedCell,
    int ContentSlideCount,
    bool RequireNames);

/// <summary>Order-independent summary of all slide text, so libraries that enumerate shapes differently compare equally.</summary>
public readonly record struct PowerPointTextObservation(int ItemCount, int CharacterCount, ulong Checksum) {
    public static PowerPointTextObservation Empty => new(0, 0, 0);

    public PowerPointTextObservation Add(string text) {
        ulong hash = 14695981039346656037UL;
        int characters = 0;
        foreach (char character in text) {
            if (character is '\r' or '\n' or '\v') continue;
            hash ^= character;
            hash *= 1099511628211UL;
            characters++;
        }
        if (characters == 0) return this;
        return new PowerPointTextObservation(ItemCount + 1, CharacterCount + characters, Checksum + hash);
    }

    internal static PowerPointTextObservation Expected(int slideCount) {
        PowerPointTextObservation observation = Empty;
        for (int slide = 0; slide < slideCount; slide++) {
            observation = observation.Add(PowerPointComparisonCorpus.Title(slide));
            observation = observation.Add(string.Concat(Enumerable.Range(0, PowerPointComparisonCorpus.BodyLineCount)
                .Select(line => PowerPointComparisonCorpus.BodyLine(slide, line))));
            for (int row = 0; row < PowerPointComparisonCorpus.TableRows; row++) {
                for (int column = 0; column < PowerPointComparisonCorpus.TableColumns; column++) {
                    observation = observation.Add(PowerPointComparisonCorpus.Cell(slide, row, column));
                }
            }
        }
        return observation;
    }
}

internal static class PowerPointComparisonValidation {
    internal static void RunAll() {
        foreach (int slideCount in new[] { 20, 200 }) {
            new PowerPointCreateComparisonBenchmarks { SlideCount = slideCount }.Validate();
            new PowerPointReadComparisonBenchmarks { SlideCount = slideCount }.Setup();
            new PowerPointEditComparisonBenchmarks { SlideCount = slideCount }.Setup();
            new PowerPointCopySlidesComparisonBenchmarks { SlideCount = slideCount }.Setup();
        }
    }
}
