using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAcroFormReviewRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AppendOnlyFill_RejectsPushButtons(bool useTryFill) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Push button append guard")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "calculate",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Calculate"
        })).ToBytes();
        var values = new Dictionary<string, string> { ["calculate"] = "Off" };

        if (useTryFill) {
            PdfOperationResult<PdfDocument> result = PdfDocument.Open(authored).Forms.TryFill(values);
            Assert.False(result.Succeeded);
            Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Contains("Push-button", StringComparison.Ordinal));
        } else {
            Assert.Throws<ArgumentException>(() => PdfDocument.Open(authored).Forms.AppendRevision(values));
        }
    }

    [Fact]
    public void Create_RejectsChildBelowInheritedTerminalFieldWithWidgetKids() {
        PdfDocument document = PdfDocument.Open(BuildInheritedTerminalFieldPdf());

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "section.existing.child",
            Value = "new"
        })));

        Assert.Contains("terminal field", exception.Message, StringComparison.OrdinalIgnoreCase);
        PdfFormField existing = Assert.Single(document.Inspect().FormFields);
        Assert.Equal("section.existing", existing.Name);
        Assert.Equal("before", existing.Value);
    }

    [Fact]
    public void RewritePreservation_DetectsWidgetActionTriggerChanges() {
        byte[] original = BuildWidgetUriActionPdf("U");
        byte[] rewritten = BuildWidgetUriActionPdf("D");
        var options = new PdfRewritePreservationOptions {
            PreserveFormWidgetActions = true
        };

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(original, rewritten, options);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, static issue => issue.Feature == "FormWidgetActions");
    }

    [Fact]
    public void Move_RebuildsPushButtonAppearanceWhenFlagIsInherited() {
        byte[] source = BuildInheritedPushButtonPdf();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit =>
            edit.Move("group.run", pageNumber: 1, x: 40, y: 80, width: 180, height: 40));

        PdfFormField field = Assert.Single(result.Fields);
        Assert.True(field.IsPushButton);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfDictionary widgetDictionary = Assert.IsType<PdfDictionary>(objects[widget.ObjectNumber!.Value].Value);
        PdfDictionary appearances = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, widgetDictionary.Items["AP"]));
        PdfStream normal = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, appearances.Items["N"]));
        PdfArray boundingBox = Assert.IsType<PdfArray>(normal.Dictionary.Items["BBox"]);

        Assert.Equal(new[] { 0D, 0D, 180D, 40D }, boundingBox.Items.Cast<PdfNumber>().Select(number => number.Value));
    }

    [Fact]
    public void RewritePreservation_ComparesPageActionContentsWhenPageCountsDiffer() {
        byte[] original = BuildPageActionPdf(pageCount: 2, "https://before.example/");
        byte[] rewritten = BuildPageActionPdf(pageCount: 1, "https://after.example/");
        var options = new PdfRewritePreservationOptions {
            PreservePageCount = false,
            PreservePageGeometry = false,
            PreserveDocumentVersionState = false,
            PreserveRevisionStructure = false
        };

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(original, rewritten, options);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, static issue => issue.Feature == "PageActions");
    }

    [Fact]
    public void WidgetOwnedActiveContentTraversalInspectsIndirectMarkerNames() {
        byte[] source = BuildWidgetWithIndirectActiveMarkerPdf();

        PdfReadDocument readDocument = PdfReadDocument.Open(source);

        Assert.False(readDocument.HasOnlyWidgetOwnedActiveContent());
        Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Open(source).Forms.Edit(edit => edit.Rename("run", "renamed")));
    }

    [Fact]
    public void Create_RaisesHeaderForOpenTypeCffPushButtonAppearance() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath is null) return;
        var appearanceOptions = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO CFF", fontPath);

        PdfAcroFormEditResult result = PdfDocument.Open(BuildSinglePagePdf("1.4")).Forms.Edit(
            edit => edit.Create(new PdfFormFieldCreateOptions {
                Name = "run",
                Kind = PdfFormFieldCreationKind.PushButton,
                Caption = "Office"
            }),
            appearanceOptions);

        string raw = PdfEncoding.Latin1GetString(result.ToBytes());
        Assert.StartsWith("%PDF-1.6", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", raw, StringComparison.Ordinal);
    }

    private static byte[] BuildInheritedTerminalFieldPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [8 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Tx /T (section) /Kids [7 0 R] >>", "endobj",
            "7 0 obj", "<< /Parent 6 0 R /T (existing) /V (before) /Kids [8 0 R] >>", "endobj",
            "8 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 160 48] /P 3 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }));
    }

    private static byte[] BuildWidgetUriActionPdf(string trigger) {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [20 20 160 48] /P 3 0 R /AA << /" + trigger + " << /S /URI /URI (https://example.com) >> >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
    }

    private static byte[] BuildInheritedPushButtonPdf() {
        const string appearance = "BT /F1 10 Tf (Run) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [8 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Btn /Ff 65536 /T (group) /Kids [7 0 R] >>", "endobj",
            "7 0 obj", "<< /Parent 6 0 R /T (run) /Kids [8 0 R] >>", "endobj",
            "8 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 120 40] /P 3 0 R /MK << /CA (Run) >> /AP << /N 9 0 R >> >>", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Length " + appearance.Length + " >>", "stream", appearance, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }));
    }

    private static byte[] BuildPageActionPdf(int pageCount, string uri) {
        var lines = new List<string> {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count " + pageCount + " /Kids [3 0 R" + (pageCount == 2 ? " 4 0 R" : string.Empty) + "] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /AA << /O << /S /URI /URI (" + uri + ") >> >> >>", "endobj"
        };
        if (pageCount == 2) {
            lines.AddRange(new[] { "4 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>", "endobj" });
        }
        lines.AddRange(new[] { "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF" });
        return Encoding.ASCII.GetBytes(string.Join("\n", lines));
    }

    private static byte[] BuildWidgetWithIndirectActiveMarkerPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Tx /T (run) /Rect [20 20 160 48] /P 3 0 R /A 7 0 R /OfficeIMO << /S 9 0 R >> >>", "endobj",
            "7 0 obj", "<< /S /URI /URI (https://example.test/) >>", "endobj",
            "9 0 obj", "/Launch", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSinglePagePdf(string version) {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-" + version,
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 4 >>", "%%EOF"
        }));
    }
}
