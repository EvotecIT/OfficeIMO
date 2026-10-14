using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormOcrReviewTests {
    [Theory]
    [InlineData("reportlab-scanned-form.pdf")]
    [InlineData("reportlab-rotated-form.pdf")]
    public async Task IndependentProducerFieldsBindGeometryAndReviewedCorrectionsToExactSource(string name) {
        var document = Fixture(name);
        byte[] original = document.ToBytes();
        var review = await document.PrepareFormOcrAsync(Engine(document));
        Assert.Equal(7, review.Proposals.Count);
        var amount = review.Proposals.Single(item => item.Field.Name == "Amount");
        Assert.Equal(3, amount.Field.Actions.Count);
        Assert.Empty(amount.Field.Widgets[0].Actions);
        Assert.True(amount.CanAccept);
        Assert.Equal("123.45", amount.SuggestedValue);
        Assert.False(review.Assess(amount, "123.45").HasErrors);
        Assert.True(review.Assess(amount, "123.456").HasErrors);
        Assert.True(review.Assess(amount, "1001").HasErrors);
        Assert.True(review.Assess(amount, "not a number").HasErrors);
        var country = review.Proposals.Single(item => item.Field.Name == "Country");
        var code = review.Proposals.Single(item => item.Field.Name == "SerialCode");
        Assert.True(code.HasLowConfidence);
        Assert.True(review.Assess(code, "TOOLONG").HasErrors);
        Assert.False(review.Proposals.Single(item => item.Field.Name == "ReadOnly").CanAccept);
        Assert.False(review.Proposals.Single(item => item.Field.Name == "CustomRule").CanAccept);
        var reference = review.Proposals.Single(item => item.Field.Name == "Reference");
        var layout = document.GetPageLayouts()[1];
        var widget = reference.Field.Widgets.Single();
        var expected = layout.MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
        Assert.Equal(expected.Left, reference.Evidence[0].WidgetBounds.Left, 6);
        Assert.Equal(expected.Top, reference.Evidence[0].WidgetBounds.Top, 6);
        var accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
            [amount] = "200.50", [country] = "Germany", [code] = "Z9Y8X7", [reference] = "INV-2049"
        };
        var result = review.Apply(document, accepted);
        Assert.Equal("DE", result.Inspect().FormFieldsByName["Country"].Value);
        Assert.Equal("Z9Y8X7", result.Inspect().FormFieldsByName["SerialCode"].Value);
        Assert.Equal("200.50", result.Inspect().FormFieldsByName["Amount"].Value);
        Assert.Equal("INV-2049", result.Inspect().FormFieldsByName["Reference"].Value);
        Assert.Equal("KEEP", result.Inspect().FormFieldsByName["ReadOnly"].Value);
        Assert.Equal(amount.Field.Actions.Select(action => action.JavaScript), result.Inspect().FormFieldsByName["Amount"].Actions.Select(action => action.JavaScript));
        Assert.Equal(review.Proposals.Single(item => item.Field.Name == "CustomRule").Field.JavaScript, result.Inspect().FormFieldsByName["CustomRule"].JavaScript);
        Assert.Equal(original, document.ToBytes());
        Assert.Throws<InvalidOperationException>(() => review.Apply(result, accepted));
        var other = await document.PrepareFormOcrAsync(Engine(document));
        Assert.Throws<ArgumentException>(() => review.Apply(document, new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> { [other.Proposals[0]] = "foreign" }));
        accepted[amount] = "2000";
        Assert.Throws<ArgumentException>(() => review.Apply(document, accepted));
        Assert.ThrowsAny<OperationCanceledException>(() => review.Apply(document, accepted, new CancellationToken(true)));
        Assert.Equal(original, review.Apply(document, new Dictionary<PdfFormOcrProposal, PdfFormFieldValue>()).ToBytes());
        Assert.Equal(original, document.ToBytes());
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_FORM_OCR_OUTPUT");
        if (!string.IsNullOrWhiteSpace(evidence)) { System.IO.Directory.CreateDirectory(evidence); result.Save(System.IO.Path.Combine(evidence, name.Replace(".pdf", "-reviewed.pdf"))); }
    }

    [Fact]
    public async Task OverlappingFieldsRemainAmbiguousAndRequireExplicitValues() {
        var source = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300))).Forms.Edit(edit => {
            edit.Create(new() { Name = "First", X = 20, Y = 200, Width = 150, Height = 30 });
            edit.Create(new() { Name = "Second", X = 20, Y = 200, Width = 150, Height = 30 });
        }).ToDocument();
        var engine = new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult {
            Spans = [Word("Ambiguous", 30, 76, 80, 14, 0.99)]
        }));
        var review = await source.PrepareFormOcrAsync(engine);
        Assert.Equal(2, review.Proposals.Count);
        Assert.All(review.Proposals, proposal => Assert.True(proposal.IsAmbiguous));
        var accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> { [review.Proposals[1]] = "Reviewed" };
        var result = review.Apply(source, accepted);
        Assert.Equal(string.Empty, result.Inspect().FormFieldsByName["First"].Value);
        Assert.Equal("Reviewed", result.Inspect().FormFieldsByName["Second"].Value);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, true)]
    public async Task CertificationPlanningRemainsTheAuthorityForReviewedFormValues(int permission, bool allowed) {
        byte[] original = PdfITextInspiredCoverageTests.BuildDocMdpFormPdf(permissionLevel: permission);
        var source = PdfDocument.Load(original);
        var field = source.Inspect().FormFields.Single(item => item.Name == "Name");
        var widget = field.Widgets.Single();
        var bounds = source.GetPageLayouts()[widget.PageNumber!.Value - 1].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
        var engine = new DelegateOcrEngine("certification-fixture", (_, _) => Task.FromResult(new OcrResult {
            Spans = [Word("Reviewed", bounds.Left + 1, bounds.Top + 1, bounds.Width - 2, bounds.Height - 2, 0.99)]
        }));
        var review = await source.PrepareFormOcrAsync(engine);
        var proposal = review.Proposals.Single(item => item.Field.Name == "Name");
        var accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> { [proposal] = "Reviewed" };
        if (allowed) {
            byte[] updated = review.Apply(source, accepted).ToBytes();
            Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
            Assert.Equal("Reviewed", PdfDocument.Load(updated).Inspect().FormFieldsByName["Name"].Value);
        } else Assert.Throws<PdfMutationBlockedException>(() => review.Apply(source, accepted));
        Assert.Equal(original, source.ToBytes());
    }

    [Fact]
    public async Task LaterOptionAndCallerDocumentChangesCannotRetargetCapturedReview() {
        var source = Fixture("reportlab-scanned-form.pdf");
        var options = new PdfOcrMergeOptions { Dpi = 72, MaxPixelsPerPage = 10_000_000 };
        var review = await source.PrepareFormOcrAsync(Engine(source), options);
        var original = review.RenderPage(1);
        options.Dpi = 10000; options.MaxPixelsPerPage = 1;
        options.ReadOptions.LayoutOptions.ReadingDirection = PdfReadingDirection.RightToLeft;
        source = source.Forms.Fill(new Dictionary<string, PdfFormFieldValue> { ["FullName"] = "Different" });
        var captured = review.RenderPage(1);
        Assert.Equal(original.Width, captured.Width); Assert.Equal(original.Height, captured.Height);
        Assert.Equal(original.Bytes, captured.Bytes);
        Assert.Equal("Alex Morgan", review.Proposals.Single(item => item.Field.Name == "FullName").SuggestedValue);
        Assert.Throws<InvalidOperationException>(() => review.Apply(source, new Dictionary<PdfFormOcrProposal, PdfFormFieldValue>()));
    }

    [Theory]
    [InlineData("open", false)]
    [InlineData("owner", true)]
    public async Task RecognitionDoesNotAuthorizeFillingEncryptedSource(string password, bool allowed) {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.CopyContents
        };
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .TextField("Name", width: 180, height: 24).ToBytes();
        var source = PdfDocument.Load(bytes, new PdfLoadOptions { Password = password });
        var widget = source.Inspect().FormFieldsByName["Name"].Widgets.Single();
        var bounds = source.GetPageLayouts()[0].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
        var engine = new DelegateOcrEngine("permission-fixture", (_, _) => Task.FromResult(new OcrResult {
            Spans = [Word("Reviewed", bounds.Left + 1, bounds.Top + 1, bounds.Width - 2, bounds.Height - 2, 0.99)]
        }));
        var review = await source.PrepareFormOcrAsync(engine);
        var accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> { [review.Proposals.Single()] = "Reviewed" };
        if (allowed) {
            var result = review.Apply(source, accepted);
            Assert.True(result.Inspect().Security.HasEncryption);
            Assert.Equal("Reviewed", result.Inspect().FormFieldsByName["Name"].Value);
        } else Assert.Throws<PdfMutationBlockedException>(() => review.Apply(source, accepted));
        Assert.Equal(bytes, source.ToBytes());
    }

    internal static PdfDocument Fixture(string name) => PdfDocument.Load(System.IO.Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "PdfFormOcr", name));
    internal static IOcrEngine Engine(PdfDocument document) {
        var fields = document.Inspect().FormFields;
        var layouts = document.GetPageLayouts().ToDictionary(page => page.PageNumber);
        var text = new Dictionary<string, string> { ["FullName"] = "Alex Morgan", ["SerialCode"] = "A1B2C3", ["Country"] = "Poland",
            ["Amount"] = "123.45", ["Reference"] = "INV-2048", ["ReadOnly"] = "KEEP", ["CustomRule"] = "77" };
        return new DelegateOcrEngine("geometry-fixture", (request, token) => {
            token.ThrowIfCancellationRequested();
            var spans = fields.SelectMany(field => field.Widgets.Where(widget => widget.PageNumber == request.PageNumber).Select(widget => {
                var bounds = layouts[widget.PageNumber!.Value].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
                return Word(text[field.Name!], bounds.Left + 2, bounds.Top + 2, bounds.Right - bounds.Left - 4,
                    bounds.Bottom - bounds.Top - 4, field.Name == "SerialCode" ? 0.1 : 0.99);
            })).ToArray();
            return Task.FromResult(new OcrResult { Provider = "synthetic geometry provider", Spans = spans });
        });
    }
    private static OcrTextSpan Word(string text, double x, double y, double width, double height, double confidence) => new() {
        Text = text, Confidence = confidence, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new() { X = x, Y = y, Width = width, Height = height }
    };
}
