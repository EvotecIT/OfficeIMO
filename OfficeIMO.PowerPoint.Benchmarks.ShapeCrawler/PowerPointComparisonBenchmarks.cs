using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using OfficeIMO.PowerPoint;
using ShapeCrawler;
using Length = OfficeIMO.OfficeLength;

namespace OfficeIMO.PowerPoint.Benchmarks.Comparison;

/// <summary>Compares creating a deck where every slide has a bold title, a three-paragraph body, a 4x3 table, and a PNG picture.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "Create")]
public class PowerPointCreateComparisonBenchmarks {
    [Params(20, 200)]
    public int SlideCount { get; set; }

    [GlobalSetup]
    public void Validate() {
        // Each deck is reopened with the Open XML SDK: slide count, bold title, three-paragraph body, 4x3 table text and
        // the exact picture bytes on every slide, then the Open XML SDK validator.
        PresentationExpectation expected = PowerPointComparisonCorpus.Created(SlideCount);
        string detail = $"SlideCount={SlideCount}; {SlideCount} slides with title, body, table text and picture bytes as written; passes the Open XML SDK validator";
        PowerPointComparisonCorpus.Validate("OfficeIMO", nameof(OfficeIMO), OfficeIMO(), expected, detail);
        PowerPointComparisonCorpus.Validate("ShapeCrawler", nameof(ShapeCrawler), ShapeCrawler(), expected, detail);
        PowerPointComparisonCorpus.Validate("Open XML SDK", nameof(OpenXmlSdk), OpenXmlSdk(), expected, detail);
        PowerPointComparisonCorpus.Validate("OneDocument.Presentations", nameof(OneDocument), OneDocument(), expected, detail);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTextBox(PowerPointComparisonCorpus.Title(index),
                Length.Points(40), Length.Points(24), Length.Points(760), Length.Points(50));
            title.Name = PowerPointComparisonCorpus.TitleName;
            title.Bold = true;
            title.FontSize = Length.Points(28);
            PowerPointTextBox body = slide.AddTextBox(PowerPointComparisonCorpus.BodyLine(index, 0),
                Length.Points(40), Length.Points(90), Length.Points(420), Length.Points(160));
            body.Name = PowerPointComparisonCorpus.BodyName;
            for (int line = 1; line < PowerPointComparisonCorpus.BodyLineCount; line++) {
                body.AddParagraph(PowerPointComparisonCorpus.BodyLine(index, line));
            }
            PowerPointTable table = slide.AddTable(PowerPointComparisonCorpus.TableRows, PowerPointComparisonCorpus.TableColumns,
                Length.Points(500), Length.Points(90), Length.Points(360), Length.Points(120));
            for (int row = 0; row < PowerPointComparisonCorpus.TableRows; row++) {
                for (int column = 0; column < PowerPointComparisonCorpus.TableColumns; column++) {
                    table.GetCell(row, column).Text = PowerPointComparisonCorpus.Cell(index, row, column);
                }
            }
            using var image = new MemoryStream(PowerPointComparisonCorpus.Image, writable: false);
            slide.AddImage(image, global::OfficeIMO.Drawing.OfficeImageFormat.Png,
                Length.Points(40), Length.Points(300), Length.Points(230), Length.Points(54));
        }
        return presentation.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:ShapeCrawler", "case:Default")]
    public byte[] ShapeCrawler() {
        using var presentation = new Presentation();
        while (presentation.Slides.Count < SlideCount) presentation.Slides.Add(1);
        for (int index = 0; index < SlideCount; index++) {
            IUserSlideShapeCollection shapes = presentation.Slide(index + 1).Shapes;
            shapes.AddTextBox(40, 24, 760, 50, PowerPointComparisonCorpus.Title(index));
            IShape title = shapes.Last();
            title.Name = PowerPointComparisonCorpus.TitleName;
            title.SetFontSize(28);
            title.TextBox!.Paragraphs.First().Portions.First().Font!.IsBold = true;
            shapes.AddTextBox(40, 90, 420, 160, PowerPointComparisonCorpus.BodyLine(index, 0));
            IShape body = shapes.Last();
            body.Name = PowerPointComparisonCorpus.BodyName;
            for (int line = 1; line < PowerPointComparisonCorpus.BodyLineCount; line++) {
                body.TextBox!.Paragraphs.Add();
                body.TextBox.Paragraphs.Last().Text = PowerPointComparisonCorpus.BodyLine(index, line);
            }
            shapes.AddTable(500, 90, PowerPointComparisonCorpus.TableColumns, PowerPointComparisonCorpus.TableRows);
            ITable table = shapes.Last().Table!;
            for (int row = 0; row < PowerPointComparisonCorpus.TableRows; row++) {
                for (int column = 0; column < PowerPointComparisonCorpus.TableColumns; column++) {
                    table[row, column].TextBox!.SetText(PowerPointComparisonCorpus.Cell(index, row, column));
                }
            }
            using var image = new MemoryStream(PowerPointComparisonCorpus.Image, writable: false);
            shapes.AddPicture(image);
            IShape picture = shapes.Last();
            picture.X = 40;
            picture.Y = 300;
            picture.Width = 230;
            picture.Height = 54;
        }
        using var output = new MemoryStream();
        presentation.Save(output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.Create(SlideCount);

    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Presentations", "case:Default")]
    public byte[] OneDocument() => PowerPointOneDocumentLane.Create(SlideCount);
}

/// <summary>Base for workloads that start from the same Open XML SDK-authored deck.</summary>
public abstract class PowerPointFixtureBenchmarks {
    protected byte[] Fixture = Array.Empty<byte>();

    [Params(20, 200)]
    public int SlideCount { get; set; }

    protected void CreateFixture() {
        Fixture = PowerPointOpenXmlSdkLane.Create(SlideCount);
        PowerPointComparisonCorpus.Validate("fixture", Fixture, PowerPointComparisonCorpus.Created(SlideCount));
    }
}

/// <summary>Compares loading a deck and reading the text of every text box and table cell.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "Read")]
public class PowerPointReadComparisonBenchmarks : PowerPointFixtureBenchmarks {
    private PowerPointTextObservation _expected;

    [GlobalSetup]
    public void Setup() {
        CreateFixture();
        _expected = PowerPointTextObservation.Expected(SlideCount);
        Ensure("OfficeIMO", nameof(OfficeIMO), OfficeIMO());
        Ensure("ShapeCrawler", nameof(ShapeCrawler), ShapeCrawler());
        Ensure("Open XML SDK", nameof(OpenXmlSdk), OpenXmlSdk());
        Ensure("OneDocument.Presentations", nameof(OneDocument), OneDocument());
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public PowerPointTextObservation OfficeIMO() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        PowerPointTextObservation observation = PowerPointTextObservation.Empty;
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointTextBox textBox in slide.TextBoxes) observation = observation.Add(textBox.Text);
            foreach (PowerPointTable table in slide.Tables) {
                int rows = table.Rows.Count;
                int columns = table.Columns.Count;
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        observation = observation.Add(table.GetCell(row, column).Text);
                    }
                }
            }
        }
        return observation;
    }

    [Benchmark]
    [BenchmarkCategory("library:ShapeCrawler", "case:Default")]
    public PowerPointTextObservation ShapeCrawler() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using var presentation = new Presentation(stream);
        PowerPointTextObservation observation = PowerPointTextObservation.Empty;
        foreach (IUserSlide slide in presentation.Slides) {
            foreach (IShape shape in slide.Shapes) {
                if (shape.Table is { } table) {
                    for (int row = 0; row < table.Rows.Count; row++) {
                        for (int column = 0; column < table.Columns.Count; column++) {
                            observation = observation.Add(table[row, column].TextBox?.Text ?? string.Empty);
                        }
                    }
                } else if (shape.TextBox is { } textBox) {
                    observation = observation.Add(textBox.Text);
                }
            }
        }
        return observation;
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public PowerPointTextObservation OpenXmlSdk() => PowerPointOpenXmlSdkLane.Read(Fixture);

    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Presentations", "case:Default")]
    public PowerPointTextObservation OneDocument() => PowerPointOneDocumentLane.Read(Fixture);

    /// <summary>Requires every text box and table cell text of the deck, in order (count, characters and checksum).</summary>
    private void Ensure(string library, string lane, PowerPointTextObservation actual) {
        string detail = $"SlideCount={SlideCount}; observed {_expected} as written";
        if (actual != _expected) {
            BenchmarkOutputValidationLog.Failed(library, detail + ": observed " + actual, "Default");
            throw new InvalidDataException(lane + " observed " + actual + "; expected " + _expected + ".");
        }
        BenchmarkOutputValidationLog.Passed(library, detail, "Default");
    }
}

/// <summary>Compares loading a deck, changing every title and one table cell per slide, and saving.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "Edit")]
public class PowerPointEditComparisonBenchmarks : PowerPointFixtureBenchmarks {
    [GlobalSetup]
    public void Setup() {
        CreateFixture();
        // Each saved deck is reopened with the Open XML SDK: every slide must carry its edited title and edited cell
        // with the rest of its content and picture unchanged, then pass the Open XML SDK validator.
        PresentationExpectation expected = PowerPointComparisonCorpus.Edited(SlideCount);
        string detail = $"SlideCount={SlideCount}; {SlideCount} slides with the edited title and cell and unchanged content; passes the Open XML SDK validator";
        PowerPointComparisonCorpus.Validate("OfficeIMO", nameof(OfficeIMO), OfficeIMO(), expected, detail);
        PowerPointComparisonCorpus.Validate("ShapeCrawler", nameof(ShapeCrawler), ShapeCrawler(), expected, detail);
        PowerPointComparisonCorpus.Validate("Open XML SDK", nameof(OpenXmlSdk), OpenXmlSdk(), expected, detail);
        PowerPointComparisonCorpus.Validate("OneDocument.Presentations", nameof(OneDocument), OneDocument(), expected, detail);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        int index = 0;
        foreach (PowerPointSlide slide in presentation.Slides) {
            slide.TextBoxes.First(textBox => textBox.Name == PowerPointComparisonCorpus.TitleName).Text =
                PowerPointComparisonCorpus.EditedTitle(index++);
            slide.Tables.First().GetCell(1, 1).Text = PowerPointComparisonCorpus.EditedCell;
        }
        return presentation.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:ShapeCrawler", "case:Default")]
    public byte[] ShapeCrawler() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using var presentation = new Presentation(stream);
        for (int index = 0; index < presentation.Slides.Count; index++) {
            IUserSlide slide = presentation.Slide(index + 1);
            slide.Shape(PowerPointComparisonCorpus.TitleName).TextBox!.SetText(PowerPointComparisonCorpus.EditedTitle(index));
            slide.Shape(PowerPointComparisonCorpus.TableName).Table![1, 1].TextBox!.SetText(PowerPointComparisonCorpus.EditedCell);
        }
        using var output = new MemoryStream();
        presentation.Save(output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.Edit(Fixture);

    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Presentations", "case:Default")]
    public byte[] OneDocument() => PowerPointOneDocumentLane.Edit(Fixture);
}

/// <summary>Compares loading a deck, appending a copy of every slide, and saving.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "CopySlides")]
public class PowerPointCopySlidesComparisonBenchmarks : PowerPointFixtureBenchmarks {
    [GlobalSetup]
    public void Setup() {
        CreateFixture();
        // Each saved deck is reopened with the Open XML SDK: twice the slides, each copy with its source slide's
        // content and picture bytes, then the Open XML SDK validator.
        PresentationExpectation expected = PowerPointComparisonCorpus.Copied(SlideCount);
        string detail = $"SlideCount={SlideCount}; {SlideCount * 2} slides, each copy matching its source; passes the Open XML SDK validator";
        PowerPointComparisonCorpus.Validate("OfficeIMO", nameof(OfficeIMO), OfficeIMO(), expected, detail);
        PowerPointComparisonCorpus.Validate("ShapeCrawler", nameof(ShapeCrawler), ShapeCrawler(), expected, detail);
        PowerPointComparisonCorpus.Validate("Open XML SDK", nameof(OpenXmlSdk), OpenXmlSdk(), expected, detail);
        PowerPointComparisonCorpus.Validate("OneDocument.Presentations", nameof(OneDocument), OneDocument(), expected, detail);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        int count = presentation.Slides.Count;
        for (int index = 0; index < count; index++) presentation.DuplicateSlide(index, presentation.Slides.Count);
        return presentation.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:ShapeCrawler", "case:Default")]
    public byte[] ShapeCrawler() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using var presentation = new Presentation(stream);
        int count = presentation.Slides.Count;
        for (int index = 1; index <= count; index++) presentation.Slides.Add(presentation.Slide(index));
        using var output = new MemoryStream();
        presentation.Save(output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.CopySlides(Fixture);

    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Presentations", "case:Default")]
    public byte[] OneDocument() => PowerPointOneDocumentLane.CopySlides(Fixture);
}
