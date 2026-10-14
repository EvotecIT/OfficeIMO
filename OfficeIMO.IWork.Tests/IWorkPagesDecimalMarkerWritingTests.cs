using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesDecimalMarkerWritingTests {
    [Theory]
    [InlineData("plain")]
    [InlineData("mixed-font")]
    [InlineData("mixed-size")]
    public void Decimal_punctuation_distinguishes_document_identity_with_deterministic_unused_group_handling(string mode) {
        var identifiers = new List<string>();
        foreach (var style in Enum.GetValues<IWorkDecimalListMarkerStyle>()) {
            var document = Create(style);
            byte[] bytes = document.SaveBytes();
            Assert.Equal(bytes, Create(style).SaveBytes());
            using var package = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
            using var identifier = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
            identifiers.Add(identifier.ReadToEnd());
            document.CreateNumberedList(style == IWorkDecimalListMarkerStyle.Period
                ? IWorkDecimalListMarkerStyle.Parentheses : IWorkDecimalListMarkerStyle.Period);
            Assert.Equal(bytes, document.SaveBytes());
            string label = style == IWorkDecimalListMarkerStyle.Period ? "1."
                : style == IWorkDecimalListMarkerStyle.Parentheses ? "(1)" : "1)";
            Assert.Equal(label, IWorkSourceDocument.Open(bytes).ReadPages().Body.Paragraphs[0].ListLabel);
        }
        Assert.Equal(3, identifiers.Distinct(StringComparer.Ordinal).Count());

        IWorkPagesDocument Create(IWorkDecimalListMarkerStyle style) {
            var document = IWorkPagesDocument.Create();
            var list = document.CreateNumberedList(style, "Arial", 12);
            if (mode == "plain") list.AddItem("Same marker body");
            else list.AddRichItem(new[] {
                new IWorkTextRunData("Same marker", "Times New Roman", mode == "mixed-size" ? 18 : 12, bold: true),
                new IWorkTextRunData(" body", "Courier New", mode == "mixed-size" ? 8 : 12, italic: true)
            });
            return document;
        }
    }

    [Theory]
    [InlineData(IWorkDecimalListMarkerStyle.Parentheses, "(1)")]
    [InlineData(IWorkDecimalListMarkerStyle.ClosingParenthesis, "1)")]
    public void Decimal_marker_creation_and_text_edits_preserve_nine_level_counters_and_saved_markers(
        IWorkDecimalListMarkerStyle style, string sample) {
        var model = IWorkPagesNineLevelWritingTests.Example(style);
        Assert.Equal(style, model.NumberedLists[0].MarkerStyle);
        Assert.Equal(IWorkDecimalListMarkerStyle.Period, model.NumberedLists[1].MarkerStyle);
        byte[] bytes = model.SaveBytes(); Assert.Equal(bytes, model.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes);
        Check(source, bytes);
        var body = source.ReadPages().Body;
        var edit = source.BeginEdit();
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 2, "ROOT SEVEN");
        byte[] changed = edit.SaveBytes(new() { AllowUnassessedVisualCaches = true });
        Check(IWorkSourceDocument.Open(changed), changed);
        Assert.Equal("Root seven", body.Paragraphs[1].Text);
        Assert.Equal("ROOT SEVEN", IWorkSourceDocument.Open(changed).ReadPages().Body.Paragraphs[1].Text);
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 2, "Root seven");
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);

        void Check(IWorkSourceDocument input, byte[] packageBytes) {
            var projection = input.ReadPages();
            Assert.True(projection.Body.IsFormattingComplete); Assert.Empty(projection.Diagnostics);
            var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
            Assert.Equal(IWorkPagesNineLevelWritingTests.Numbers, items.Select(p => p.ListNumber!.Value));
            Assert.Equal(IWorkPagesNineLevelWritingTests.Levels, items.Select(p => p.ListLevel));
            Assert.All(items.Where(p => p.Text != "Other root three"), p => Assert.Equal(sample, p.ListLabel));
            Assert.All(items, p => Assert.Equal(new string(' ', p.ListLevel * 2) + p.ListNumber + ". " + p.Text,
                IWorkReadProjection.RichTextMarkdown(p)));
            using var result = input.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            result.RequireCompleteEditableReconstruction();
            using var savedBytes = new MemoryStream(); result.Value.Save(savedBytes); savedBytes.Position = 0;
            using var saved = WordDocument.Load(savedBytes);
            string[] markers = items.Select(p => p.Text == "Other root three" ? "3."
                : style == IWorkDecimalListMarkerStyle.Parentheses ? "(" + p.ListNumber + ")" : p.ListNumber + ")").ToArray();
            var reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
            using var packageStream = new MemoryStream(packageBytes, writable: false);
            var read = reader.ReadDocument(packageStream, "markers.pages");
            Assert.Equal(markers, read.Blocks.Where(block => block.Kind == "list-item").Select(block => block.Marker));
            Assert.Equal(markers, WordDocumentTraversal.BuildListMarkers(saved).Values.Select(value => value.Marker));
            Assert.Equal(IWorkPagesNineLevelWritingTests.Numbers,
                WordDocumentTraversal.BuildListIndices(saved).Values.Select(value => value.Index));
            Assert.Empty(saved.ValidateDocument());
        }
    }

    [Fact]
    public void Unsupported_decimal_marker_styles_do_not_mutate_the_creation_model() {
        var model = IWorkPagesDocument.Create(); model.AddParagraph("Keep this body");
        byte[] original = model.SaveBytes();
        foreach (var style in new[] { (IWorkDecimalListMarkerStyle)(-1), (IWorkDecimalListMarkerStyle)3 })
            Assert.Throws<ArgumentOutOfRangeException>(() => model.CreateNumberedList(style));
        Assert.Empty(model.NumberedLists); Assert.Equal(original, model.SaveBytes());
        Assert.Equal(IWorkDecimalListMarkerStyle.Period, model.CreateNumberedList().MarkerStyle);
    }

    [Theory]
    [InlineData("parentheses", "(1)", true)]
    [InlineData("closing", "1)", false)]
    public void Apple_saved_marker_forms_retain_sparse_counters_and_native_insertions(
        string name, string sample, bool parentheses) {
        foreach (bool edited in new[] { false, true }) {
            string suffix = edited ? "edited" : "saved";
            var source = IWorkSourceDocument.Open(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
                "native-exports", "pages-decimal-" + name + "-" + suffix + "-v15.4.pages"));
            var projection = source.ReadPages(); var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
            Assert.True(projection.Body.IsFormattingComplete); Assert.Empty(projection.Diagnostics);
            Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
            var numbers = IWorkPagesNineLevelWritingTests.Numbers.Concat(edited ? new[] { 2 } : Array.Empty<int>()).ToArray();
            var levels = IWorkPagesNineLevelWritingTests.Levels.Concat(edited ? new[] { 8 } : Array.Empty<int>()).ToArray();
            Assert.Equal(numbers, items.Select(p => p.ListNumber!.Value)); Assert.Equal(levels, items.Select(p => p.ListLevel));
            Assert.All(items, p => Assert.Equal(p.Text == "Other root three" ? "1." : sample, p.ListLabel));
            if (edited) {
                Assert.Equal("Root novel", projection.Body.Paragraphs[1].Text);
                Assert.Equal("Level EIGHT", projection.Body.Paragraphs[9].Text);
                Assert.Equal("Native deep two", items[^1].Text);
            }
            using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            result.RequireCompleteEditableReconstruction(); using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
            using var saved = WordDocument.Load(bytes);
            Assert.Equal(items.Select(p => p.Text == "Other root three" ? "3."
                    : parentheses ? "(" + p.ListNumber + ")" : p.ListNumber + ")"),
                WordDocumentTraversal.BuildListMarkers(saved).Values.Select(p => p.Marker));
            Assert.Equal(numbers, WordDocumentTraversal.BuildListIndices(saved).Values.Select(p => p.Index));
            Assert.Empty(saved.ValidateDocument());
        }
    }
}
