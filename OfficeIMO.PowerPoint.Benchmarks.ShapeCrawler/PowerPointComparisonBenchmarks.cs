using BenchmarkDotNet.Attributes;
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
        PresentationExpectation expected = PowerPointComparisonCorpus.Created(SlideCount);
        PowerPointComparisonCorpus.Validate(nameof(OfficeIMO), OfficeIMO(), expected);
        PowerPointComparisonCorpus.Validate(nameof(ShapeCrawler), ShapeCrawler(), expected);
        PowerPointComparisonCorpus.Validate(nameof(OpenXmlSdk), OpenXmlSdk(), expected);
    }

    [Benchmark(Baseline = true)]
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
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.Create(SlideCount);
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
        Ensure(nameof(OfficeIMO), OfficeIMO());
        Ensure(nameof(ShapeCrawler), ShapeCrawler());
        Ensure(nameof(OpenXmlSdk), OpenXmlSdk());
    }

    [Benchmark(Baseline = true)]
    public PowerPointTextObservation OfficeIMO() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        PowerPointTextObservation observation = PowerPointTextObservation.Empty;
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointTextBox textBox in slide.TextBoxes) observation = observation.Add(textBox.Text);
            foreach (PowerPointTable table in slide.Tables) {
                int rows = table.Rows;
                int columns = table.Columns;
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
    public PowerPointTextObservation OpenXmlSdk() => PowerPointOpenXmlSdkLane.Read(Fixture);

    private void Ensure(string lane, PowerPointTextObservation actual) {
        if (actual != _expected) throw new InvalidDataException(lane + " observed " + actual + "; expected " + _expected + ".");
    }
}

/// <summary>Compares loading a deck, changing every title and one table cell per slide, and saving.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "Edit")]
public class PowerPointEditComparisonBenchmarks : PowerPointFixtureBenchmarks {
    [GlobalSetup]
    public void Setup() {
        CreateFixture();
        PresentationExpectation expected = PowerPointComparisonCorpus.Edited(SlideCount);
        PowerPointComparisonCorpus.Validate(nameof(OfficeIMO), OfficeIMO(), expected);
        PowerPointComparisonCorpus.Validate(nameof(ShapeCrawler), ShapeCrawler(), expected);
        PowerPointComparisonCorpus.Validate(nameof(OpenXmlSdk), OpenXmlSdk(), expected);
    }

    [Benchmark(Baseline = true)]
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
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.Edit(Fixture);
}

/// <summary>Compares loading a deck, appending a copy of every slide, and saving.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("PowerPoint", "Comparison", "CopySlides")]
public class PowerPointCopySlidesComparisonBenchmarks : PowerPointFixtureBenchmarks {
    [GlobalSetup]
    public void Setup() {
        CreateFixture();
        PresentationExpectation expected = PowerPointComparisonCorpus.Copied(SlideCount);
        PowerPointComparisonCorpus.Validate(nameof(OfficeIMO), OfficeIMO(), expected);
        PowerPointComparisonCorpus.Validate(nameof(ShapeCrawler), ShapeCrawler(), expected);
        PowerPointComparisonCorpus.Validate(nameof(OpenXmlSdk), OpenXmlSdk(), expected);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() {
        using var stream = new MemoryStream(Fixture, writable: false);
        using PowerPointDocument presentation = PowerPointDocument.Load(stream);
        int count = presentation.Slides.Count;
        for (int index = 0; index < count; index++) presentation.DuplicateSlide(index, presentation.Slides.Count);
        return presentation.ToBytes();
    }

    [Benchmark]
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
    public byte[] OpenXmlSdk() => PowerPointOpenXmlSdkLane.CopySlides(Fixture);
}
