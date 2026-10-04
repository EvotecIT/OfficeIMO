using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("inherit", "FF0000")]
    [InlineData("override", "0000FF")]
    [InlineData("clear", null)]
    public void Keynote_background_inheritance_and_clear_survive_saved_presentation(string mode, string? color) {
        byte[]? child = mode == "inherit" ? null : mode == "clear" ? Message() : FillColor(0, 0, 1);
        using var package = KeynoteWithBuildDeclarations(ReferenceField(1, 10),
            SlideBackgroundStyle(10, child, 11), SlideBackgroundStyle(11, FillColor(1, 0, 0)));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.False(result.IsVisualFallback);
        IWorkKeynoteSlide source = Assert.Single(result.Projection.Slides);
        Assert.True(source.HasBackgroundFill);
        Assert.Equal(color, source.BackgroundColor?.RgbHex);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(color, reopened.Slides[0].BackgroundColor);
        Assert.Empty(reopened.ValidateDocument());
        if (color == null) {
            saved.Position = 0;
            using var xml = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(saved, false);
            Assert.IsType<DocumentFormat.OpenXml.Drawing.NoFill>(Assert.Single(xml.PresentationPart!.SlideParts.Single()
                .Slide!.CommonSlideData!.Background!.BackgroundProperties!.ChildElements));
            Assert.Equal(PowerPointSlideBackgroundKind.None, reopened.Slides[0].GetBackground().Kind);
        }
        package.Position = 0;
        var reader = IWorkReaderAdapter.ReadDocument(package, "background.key", new OfficeIMO.Reader.ReaderOptions(),
            new ReaderIWorkOptions(), System.Threading.CancellationToken.None);
        Assert.Contains(reader.Diagnostics, d => d.Code == "IWORK_READER_SLIDE_BACKGROUND_OMITTED");
    }

    [Theory]
    [InlineData("gradient")]
    [InlineData("alpha")]
    [InlineData("p3")]
    [InlineData("unknown-color")]
    [InlineData("non-neutral-extra-color")]
    [InlineData("malformed-extra-color")]
    [InlineData("duplicate-extra-color")]
    [InlineData("malformed")]
    public void Keynote_unsupported_backgrounds_gate_editable_conversion_with_physical_evidence(string kind) {
        byte[] fill = kind switch {
            "gradient" => BytesField(2, Message()),
            "alpha" => FillColor(1, 0, 0, .5f),
            "p3" => FillColor(1, 0, 0, space: 2),
            "unknown-color" => ExtraColorFill(FloatField(14, 1)),
            "non-neutral-extra-color" => ExtraColorFill(FloatField(13, .5f)),
            "malformed-extra-color" => ExtraColorFill(VarintField(13, 1)),
            "duplicate-extra-color" => ExtraColorFill(FloatField(13, 1), FloatField(13, 1)),
            _ => new byte[] { 0x80 }
        };
        using var package = KeynoteWithBuildDeclarations(ReferenceField(1, 10), SlideBackgroundStyle(10, fill));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Slides[0].HasBackgroundFill);
        Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.FieldPath == "11/1");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BACKGROUND_UNSUPPORTED");
        package.Position = 0;
        using var partial = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(partial.IsVisualFallback);
        Assert.Null(partial.Value.Slides[0].BackgroundColor);
    }

    private static byte[] ExtraColorFill(params byte[][] extraFields) => BytesField(1, Message(
        new[] { VarintField(1, 1), FloatField(3, 1), FloatField(4, 0), FloatField(5, 0),
            FloatField(6, 1), VarintField(12, 1) }.Concat(extraFields).ToArray()));

    [Theory]
    [InlineData("missing")]
    [InlineData("wrong-type")]
    [InlineData("duplicate")]
    [InlineData("cycle")]
    public void Keynote_background_style_graph_failures_do_not_select_an_arbitrary_fill(string kind) {
        byte[] reference = kind == "duplicate" ? Message(ReferenceField(1, 10), ReferenceField(1, 10)) : ReferenceField(1, 10);
        byte[][] records = kind switch {
            "missing" => Array.Empty<byte[]>(),
            "wrong-type" => new[] { ArchiveRecord(10, 6004, BytesField(11, BytesField(1, FillColor(1, 0, 0)))) },
            "cycle" => new[] { SlideBackgroundStyle(10, FillColor(1, 0, 0), 10) },
            _ => new[] { SlideBackgroundStyle(10, FillColor(1, 0, 0)) }
        };
        using var package = KeynoteWithBuildDeclarations(reference, records);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Slides[0].HasBackgroundFill);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BACKGROUND_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_background_inheritance_respects_the_configured_depth() {
        using var package = KeynoteWithBuildDeclarations(ReferenceField(1, 10),
            SlideBackgroundStyle(10, null, 11), SlideBackgroundStyle(11, FillColor(1, 0, 0)));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            readOptions: new IWorkReadOptions { MaximumTextStyleInheritanceDepth = 1 }, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.Slides[0].HasBackgroundFill);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Keynote_unqualified_template_background_is_not_assumed_to_be_blank(bool hasStyle) {
        using var package = KeynoteWithBuildDeclarations(Message(ReferenceField(17, 12),
            hasStyle ? ReferenceField(1, 10) : Message()), SlideBackgroundStyle(10, null),
            ArchiveRecord(12, 5, ReferenceField(1, 11)), SlideBackgroundStyle(11, FillColor(1, 0, 0)));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.FieldPath == "17");
    }

    [Fact]
    public void Keynote_unused_background_style_does_not_gate_selected_content() {
        using var package = KeynoteWithBuildDeclarations(Message(), SlideBackgroundStyle(10, BytesField(2, Message())));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.False(result.Projection.Slides[0].HasBackgroundFill);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BACKGROUND_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_native_backgrounds_match_independent_selected_style_evidence() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Fixture("keynote-backgrounds.json")));
        foreach (JsonElement expectedSource in manifest.RootElement.GetProperty("sources").EnumerateArray()) {
            string path = Fixture(expectedSource.GetProperty("path").GetString()!);
            Assert.Equal(expectedSource.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path,
                conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            Assert.False(result.IsVisualFallback);
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            foreach (IWorkKeynoteSlide slide in result.Projection.Slides) {
                JsonElement expected = expectedSource.GetProperty("slidesIncludingTemplates").EnumerateArray()
                    .Single(item => item.GetProperty("identifier").GetUInt64() == slide.SourceIdentity!.RecordIdentifier);
                string kind = expected.GetProperty("backgroundKind").GetString()!;
                Assert.Equal(kind is "solid" or "none", slide.HasBackgroundFill);
                Assert.Equal(expected.GetProperty("rgb").GetString(), slide.BackgroundColor?.RgbHex);
                Assert.Equal(slide.BackgroundColor?.RgbHex, reopened.Slides[slide.Index - 1].BackgroundColor);
                if (kind == "unsupported") Assert.Contains(result.Report.Diagnostics,
                    d => d.Code == "IWORK_KEYNOTE_BACKGROUND_UNSUPPORTED" && d.RecordIdentifier == slide.SourceIdentity!.RecordIdentifier);
            }
        }
    }

    [Fact]
    public void Keynote_15_4_native_color_profile_survives_strict_conversion_and_saved_output() {
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(
            Fixture("native-exports/keynote-colors-v15.4.key"));
        result.Report.RequireCompleteEditableReconstruction();
        Assert.False(result.IsVisualFallback);
        Assert.Equal(new[] { "56C1FF", "FF968D" }, result.Projection.Slides.Select(slide => slide.BackgroundColor?.RgbHex));
        Assert.Contains("wrapped title keeps all of its lines in this fixed frame", result.Projection.Slides[1].TitleBox!.Content.PlainText);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(new[] { "56C1FF", "FF968D" }, reopened.Slides.Select(slide => slide.BackgroundColor));
        Assert.Empty(reopened.ValidateDocument());
    }

    private static byte[] SlideBackgroundStyle(ulong id, byte[]? fill, ulong? parent = null) =>
        ArchiveRecord(id, 9, Message(parent.HasValue ? BytesField(1, ReferenceField(3, parent.Value)) : Message(),
            fill == null ? Message() : BytesField(11, BytesField(1, fill))));
}
