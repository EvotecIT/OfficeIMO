using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NpoiDocument = NPOI.XWPF.UserModel.XWPFDocument;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using XceedDocX = Xceed.Words.NET.DocX;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Compares creating a document with one caption and one distinct inline PNG per item.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "Images")]
public class WordCreateImageComparisonBenchmarks {
    private const long EmuPerPixel = 9525L;

    [Params(10, 100)]
    public int ImageCount { get; set; }

    [GlobalSetup]
    public void Validate() {
        // Each output is reopened as a ZIP/XML package: one caption per image with the expected text and one inline
        // picture per image whose embedded part carries exactly its source PNG bytes; every lane except DocX must
        // also pass the Open XML SDK validator.
        string content = $"ImageCount={ImageCount}; {ImageCount} captions and {ImageCount} inline pictures with their exact source PNG bytes";
        string conformant = content + "; passes the Open XML SDK validator";
        byte[] officeImo = OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", conformant,
            () => WordWorkloadCorpus.ValidateImageDocument(officeImo, ImageCount));
        if (BenchmarkLibrarySelection.IsEnabled("docx")) WordBenchmarkCorpus.Check("DocX", "Default", content,
            () => WordWorkloadCorpus.ValidateImageDocument(DocX(), ImageCount, requireOpenXmlSdkConformance: false));
        byte[] npoi = NPOI();
        WordBenchmarkCorpus.Check("NPOI", "Default", conformant,
            () => WordWorkloadCorpus.ValidateImageDocument(npoi, ImageCount));
        byte[] openXmlSdk = OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Minimal", conformant + "; document and image parts only",
            () => WordWorkloadCorpus.ValidateImageDocument(openXmlSdk, ImageCount));
#if NET10_0_OR_GREATER
        byte[] oneDocument = OneDocument();
        WordBenchmarkCorpus.Check("OneDocument.Documents", "Default", conformant,
            () => WordWorkloadCorpus.ValidateImageDocument(oneDocument, ImageCount));
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using WordDocument document = WordDocument.Create();
        for (int index = 0; index < ImageCount; index++) {
            document.AddParagraph(WordWorkloadCorpus.ImageCaption(index));
            using var image = new MemoryStream(WordBenchmarkImages.Image(index), writable: false);
            document.AddParagraph().AddImage(image, "image" + index + ".png",
                global::OfficeIMO.OfficeLength.Pixels(WordBenchmarkImages.WidthPixels),
                global::OfficeIMO.OfficeLength.Pixels(WordBenchmarkImages.HeightPixels));
        }
        return document.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:DocX", "case:Default")]
    public byte[] DocX() {
        using var stream = new MemoryStream();
        using (XceedDocX document = XceedDocX.Create(stream)) {
            for (int index = 0; index < ImageCount; index++) {
                document.InsertParagraph(WordWorkloadCorpus.ImageCaption(index));
                using var image = new MemoryStream(WordBenchmarkImages.Image(index), writable: false);
                Xceed.Document.NET.Image added = document.AddImage(image);
                document.InsertParagraph().AppendPicture(
                    added.CreatePicture(WordBenchmarkImages.HeightPixels, WordBenchmarkImages.WidthPixels));
            }
            document.Save();
        }
        return stream.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:NPOI", "case:Default")]
    public byte[] NPOI() {
        using var document = new NpoiDocument();
        for (int index = 0; index < ImageCount; index++) {
            document.CreateParagraph().CreateRun().SetText(WordWorkloadCorpus.ImageCaption(index));
            using var image = new MemoryStream(WordBenchmarkImages.Image(index), writable: false);
            document.CreateParagraph().CreateRun().AddPicture(
                image,
                (int)global::NPOI.XWPF.UserModel.PictureType.PNG,
                "image" + index + ".png",
                (int)(WordBenchmarkImages.WidthPixels * EmuPerPixel),
                (int)(WordBenchmarkImages.HeightPixels * EmuPerPixel));
        }
        using var stream = new MemoryStream();
        document.Write(stream);
        return stream.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Minimal")]
    public byte[] OpenXmlSdk() {
        using var stream = new MemoryStream();
        using (WordprocessingDocument document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document, autoSave: true)) {
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            var body = new Body();
            for (int index = 0; index < ImageCount; index++) {
                body.Append(new Paragraph(new Run(new Text(WordWorkloadCorpus.ImageCaption(index)))));
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var image = new MemoryStream(WordBenchmarkImages.Image(index), writable: false)) {
                    imagePart.FeedData(image);
                }
                body.Append(new Paragraph(new Run(InlinePicture(mainPart.GetIdOfPart(imagePart), (uint)index + 1))));
            }
            mainPart.Document = new OpenXmlDocument(body);
        }
        return stream.ToArray();
    }

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:Default")]
    public byte[] OneDocument() => WordOneDocumentWorkloads.CreateImages(ImageCount);
#endif

    private static DocumentFormat.OpenXml.Wordprocessing.Drawing InlinePicture(string relationshipId, uint id) {
        long cx = WordBenchmarkImages.WidthPixels * EmuPerPixel;
        long cy = WordBenchmarkImages.HeightPixels * EmuPerPixel;
        return new DocumentFormat.OpenXml.Wordprocessing.Drawing(new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = id, Name = "Picture " + id },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties { Id = 0U, Name = "image" + id + ".png" },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })) {
            DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U
        });
    }
}

/// <summary>Compares mail-merge style template filling: load, replace ten distinct fields, and save.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "TemplateFill")]
public class WordTemplateFillComparisonBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = WordWorkloadCorpus.CreateTemplateFixture(ItemCount);
        // Each saved output is reopened as a ZIP/XML package and every paragraph must carry its template with all ten
        // fields filled; every lane except DocX must also pass the Open XML SDK validator.
        string content = $"ItemCount={ItemCount}; {ItemCount} paragraphs with every field filled";
        string conformant = content + "; passes the Open XML SDK validator";
        byte[] officeImo = OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", conformant,
            () => WordWorkloadCorpus.ValidateFilledDocument(officeImo, ItemCount));
        if (BenchmarkLibrarySelection.IsEnabled("docx")) WordBenchmarkCorpus.Check("DocX", "Default", content,
            () => WordWorkloadCorpus.ValidateFilledDocument(DocX(), ItemCount, requireOpenXmlSdkConformance: false));
        byte[] npoi = NPOI();
        WordBenchmarkCorpus.Check("NPOI", "Default", conformant,
            () => WordWorkloadCorpus.ValidateFilledDocument(npoi, ItemCount));
        byte[] openXmlSdk = OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Default", conformant,
            () => WordWorkloadCorpus.ValidateFilledDocument(openXmlSdk, ItemCount));
#if NET10_0_OR_GREATER
        byte[] oneDocument = OneDocument();
        WordBenchmarkCorpus.Check("OneDocument.Documents", "Default", conformant,
            () => WordWorkloadCorpus.ValidateFilledDocument(oneDocument, ItemCount));
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using var input = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(input);
        var matchCase = new global::OfficeIMO.OfficeTextSearchOptions { MatchCase = true };
        foreach (KeyValuePair<string, string> field in WordWorkloadCorpus.Fields) {
            document.ReplaceAll(field.Key, field.Value, matchCase);
        }
        return document.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:DocX", "case:Default")]
    public byte[] DocX() {
        using var input = new MemoryStream(_fixture, writable: false);
        using XceedDocX document = XceedDocX.Load(input);
        foreach (KeyValuePair<string, string> field in WordWorkloadCorpus.Fields) {
            document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions {
                SearchValue = field.Key,
                NewValue = field.Value,
                RegExOptions = System.Text.RegularExpressions.RegexOptions.None
            });
        }
        using var output = new MemoryStream();
        document.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>NPOI has no document-level replace; this is the usual single pass over runs.</summary>
    [Benchmark]
    [BenchmarkCategory("library:NPOI", "case:Default")]
    public byte[] NPOI() {
        using var input = new MemoryStream(_fixture, writable: false);
        using var document = new NpoiDocument(input);
        foreach (NPOI.XWPF.UserModel.XWPFParagraph paragraph in document.Paragraphs) {
            foreach (NPOI.XWPF.UserModel.XWPFRun run in paragraph.Runs) {
                string text = run.GetText(0) ?? string.Empty;
                string updated = ReplaceFields(text);
                if (!ReferenceEquals(text, updated)) run.SetText(updated, 0);
            }
        }
        using var output = new MemoryStream();
        document.Write(output);
        return output.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public byte[] OpenXmlSdk() {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(_fixture);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            MainDocumentPart mainPart = document.MainDocumentPart!;
            foreach (Text text in mainPart.Document!.Body!.Descendants<Text>()) {
                string updated = ReplaceFields(text.Text);
                if (!ReferenceEquals(text.Text, updated)) text.Text = updated;
            }
            mainPart.Document.Save();
        }
        return stream.ToArray();
    }

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:Default")]
    public byte[] OneDocument() => WordOneDocumentWorkloads.FillTemplate(_fixture);
#endif

    private static string ReplaceFields(string text) {
        if (text.IndexOf("{{", StringComparison.Ordinal) < 0) return text;
        string updated = text;
        foreach (KeyValuePair<string, string> field in WordWorkloadCorpus.Fields) {
            updated = updated.Replace(field.Key, field.Value, StringComparison.Ordinal);
        }
        return updated;
    }
}

/// <summary>Compares loading a report and reading every table cell through each library's table model.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "ReadTables")]
public class WordReadTableComparisonBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();
    private WordReadObservation _expected;

    [Params(100, 1000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = WordWorkloadCorpus.CreateReportFixture(RowCount);
        _expected = WordWorkloadCorpus.ExpectedReportCells(RowCount);
        Ensure("OfficeIMO", nameof(OfficeIMO), OfficeIMO());
        if (BenchmarkLibrarySelection.IsEnabled("docx")) Ensure("DocX", nameof(DocX), DocX());
        Ensure("NPOI", nameof(NPOI), NPOI());
        Ensure("Open XML SDK", nameof(OpenXmlSdk), OpenXmlSdk());
#if NET10_0_OR_GREATER
        Ensure("OneDocument.Documents", nameof(OneDocument), OneDocument());
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public WordReadObservation OfficeIMO() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        WordReadObservation observation = WordReadObservation.Empty;
        foreach (WordTable table in document.Tables) {
            foreach (WordTableRow row in table.Rows) {
                foreach (WordTableCell cell in row.Cells) {
                    observation = observation.Add(string.Concat(cell.Paragraphs.Select(paragraph => paragraph.Text)));
                }
            }
        }
        return observation;
    }

    [Benchmark]
    [BenchmarkCategory("library:DocX", "case:Default")]
    public WordReadObservation DocX() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using XceedDocX document = XceedDocX.Load(stream);
        WordReadObservation observation = WordReadObservation.Empty;
        foreach (Xceed.Document.NET.Table table in document.Tables) {
            foreach (Xceed.Document.NET.Row row in table.Rows) {
                foreach (Xceed.Document.NET.Cell cell in row.Cells) {
                    observation = observation.Add(string.Concat(cell.Paragraphs.Select(paragraph => paragraph.Text)));
                }
            }
        }
        return observation;
    }

    [Benchmark]
    [BenchmarkCategory("library:NPOI", "case:Default")]
    public WordReadObservation NPOI() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using var document = new NpoiDocument(stream);
        WordReadObservation observation = WordReadObservation.Empty;
        foreach (NPOI.XWPF.UserModel.XWPFTable table in document.Tables) {
            foreach (NPOI.XWPF.UserModel.XWPFTableRow row in table.Rows) {
                foreach (NPOI.XWPF.UserModel.XWPFTableCell cell in row.GetTableCells()) {
                    observation = observation.Add(cell.GetText());
                }
            }
        }
        return observation;
    }

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
    public WordReadObservation OpenXmlSdk() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        WordReadObservation observation = WordReadObservation.Empty;
        foreach (Table table in document.MainDocumentPart!.Document!.Body!.Elements<Table>()) {
            foreach (TableRow row in table.Elements<TableRow>()) {
                foreach (TableCell cell in row.Elements<TableCell>()) {
                    observation = observation.Add(cell.InnerText);
                }
            }
        }
        return observation;
    }

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:Default")]
    public WordReadObservation OneDocument() => WordOneDocumentWorkloads.ReadTables(_fixture);
#endif

    /// <summary>Requires every cell, in order, with the report text (count, characters and checksum).</summary>
    private void Ensure(string library, string implementation, WordReadObservation actual) =>
        WordBenchmarkCorpus.Check(library, "Default",
            $"RowCount={RowCount}; {_expected.ParagraphCount} cells, {_expected.CharacterCount} characters, checksum {_expected.Checksum:X16} as written",
            () => Ensure(implementation, actual));

    private void Ensure(string implementation, WordReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(implementation + " observed " + actual + "; expected " + _expected + ".");
        }
    }
}
