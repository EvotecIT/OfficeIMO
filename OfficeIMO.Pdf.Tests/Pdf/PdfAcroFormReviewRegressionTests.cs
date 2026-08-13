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
    public void Create_RejectsChildBelowTerminalFieldThatOnlyInheritsItsType() {
        PdfDocument document = PdfDocument.Open(BuildInheritedTerminalFieldWithoutKidsPdf());

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
        byte[] source = BuildInheritedPushButtonPdf(includeInteractiveAppearances: false);

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
    public void MoveThenRenamePreservesQueuedInheritedFieldValue() {
        byte[] source = BuildInheritedTerminalFieldPdf();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .Move("section.existing", pageNumber: 1, x: 40, y: 80, width: 180, height: 40)
            .Rename("section.existing", "section.renamed"));

        PdfFormField field = Assert.Single(result.Fields);
        Assert.Equal("section.renamed", field.Name);
        Assert.Equal("before", field.Value);
    }

    [Fact]
    public void Edit_UsesLastDefaultValueAssignedInTransaction() {
        byte[] source = PdfDocument.Create().TextField("name", value: "Ada").ToBytes();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit => edit
            .SetDefaultValue("name", "first")
            .SetDefaultValue("name", "second"));

        Assert.Equal("second", Assert.Single(result.Fields).DefaultValue);
    }

    [Fact]
    public void Move_PreservesNonzeroWidgetGenerationInPageAnnotations() {
        byte[] source = BuildNonzeroGenerationWidgetPdf();
        Assert.Equal(2, PdfSyntax.ParseObjects(source, null).Map[6].Generation);

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit =>
            edit.Move("name", pageNumber: 1, x: 40, y: 80, width: 180, height: 40));
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfDictionary page = Assert.IsType<PdfDictionary>(objects[3].Value);
        PdfArray annotations = Assert.IsType<PdfArray>(page.Items["Annots"]);
        PdfReference widgetReference = Assert.IsType<PdfReference>(Assert.Single(annotations.Items));

        Assert.Equal(objects[widgetReference.ObjectNumber].Generation, widgetReference.Generation);
    }

    [Fact]
    public void Move_RejectsResizingPushButtonWithRolloverAndDownAppearances() {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(BuildInheritedPushButtonPdf()).Forms.Edit(edit =>
                edit.Move("group.run", pageNumber: 1, x: 40, y: 80, width: 180, height: 40)));

        Assert.Contains("rollover or down", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Move_PreservesCustomPushButtonNormalAppearanceWhenOnlyPositionChanges() {
        PdfAcroFormEditResult result = PdfDocument.Open(BuildInheritedPushButtonPdf()).Forms.Edit(edit =>
            edit.Move("group.run", pageNumber: 1, x: 80, y: 120, width: 100, height: 20));
        PdfFormWidget widget = Assert.Single(Assert.Single(result.Fields).Widgets);
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfDictionary widgetDictionary = Assert.IsType<PdfDictionary>(objects[widget.ObjectNumber!.Value].Value);
        PdfDictionary appearances = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, widgetDictionary.Items["AP"]));
        PdfStream normal = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, appearances.Items["N"]));

        Assert.Contains("(Run) Tj", PdfEncoding.Latin1GetString(normal.Data), StringComparison.Ordinal);
        Assert.Equal(new[] { 0D, 0D, 100D, 20D }, Assert.IsType<PdfArray>(normal.Dictionary.Items["BBox"]).Items.Cast<PdfNumber>().Select(number => number.Value));
    }

    [Fact]
    public void Move_DetachesSharedPushButtonAppearanceDictionary() {
        byte[] source = BuildSharedPushButtonAppearancePdf();

        PdfAcroFormEditResult result = PdfDocument.Open(source).Forms.Edit(edit =>
            edit.Move("first", pageNumber: 1, x: 40, y: 80, width: 180, height: 40));
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result.ToBytes(), null).Map;
        PdfDictionary first = RequireNamedField(objects, "first");
        PdfDictionary second = RequireNamedField(objects, "second");
        PdfDictionary firstAppearances = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, first.Items["AP"]));
        PdfDictionary secondAppearances = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, second.Items["AP"]));
        PdfStream firstNormal = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, firstAppearances.Items["N"]));
        PdfStream secondNormal = Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, secondAppearances.Items["N"]));

        Assert.Equal(new[] { 0D, 0D, 180D, 40D }, Assert.IsType<PdfArray>(firstNormal.Dictionary.Items["BBox"]).Items.Cast<PdfNumber>().Select(number => number.Value));
        Assert.Equal(new[] { 0D, 0D, 100D, 20D }, Assert.IsType<PdfArray>(secondNormal.Dictionary.Items["BBox"]).Items.Cast<PdfNumber>().Select(number => number.Value));
    }

    [Fact]
    public void Edit_RejectsClearingPushButtonSemanticFlag() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Push button flags")).ToBytes();
        byte[] authored = PdfDocument.Open(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "run",
            Kind = PdfFormFieldCreationKind.PushButton,
            Caption = "Run"
        })).ToBytes();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(authored).Forms.Edit(edit => edit.SetFlags("run", 0)));

        Assert.Contains("push-button flag", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Edit_RejectsConvertingCheckBoxToPushButton() {
        byte[] source = PdfDocument.Create().CheckBox("Action", isChecked: true).ToBytes();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Open(source).Forms.Edit(edit => edit.SetFlags("Action", 1 << 16)));

        Assert.Contains("Converting", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AttachmentEdit_PreservesHeaderWhenCatalogAlreadyDeclaresOpenTypeVersion() {
        byte[] source = BuildCatalogVersionedOpenTypePdf();

        byte[] output = PdfAttachmentEditor.Add(source, new PdfEmbeddedFile("note.txt", Encoding.UTF8.GetBytes("note"))).ToBytes();

        Assert.StartsWith("%PDF-1.4", PdfEncoding.Latin1GetString(output), StringComparison.Ordinal);
        Assert.Equal("1.6", PdfInspector.Inspect(output).CatalogVersion);
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(output));
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
    public void RewritePreservation_NormalizesPageReferencesInsideActionDestinations() {
        byte[] original = BuildPageDestinationActionPdf(firstPageObjectNumber: 3, secondPageObjectNumber: 4);
        byte[] rewritten = BuildPageDestinationActionPdf(firstPageObjectNumber: 8, secondPageObjectNumber: 9);

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(original, rewritten);

        Assert.True(report.IsPreserved);
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

    [Fact]
    public void Create_RaisesAnOverridingCatalogVersionForOpenTypeCffPushButtonAppearance() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath is null) return;
        var appearanceOptions = new PdfFormFillerOptions()
            .UseAppearanceFontFile("OfficeIMO CFF", fontPath);

        PdfAcroFormEditResult result = PdfDocument.Open(BuildSinglePagePdf("1.4", "1.4")).Forms.Edit(
            edit => edit.Create(new PdfFormFieldCreateOptions {
                Name = "run",
                Kind = PdfFormFieldCreationKind.PushButton,
                Caption = "Office"
            }),
            appearanceOptions);

        string raw = PdfEncoding.Latin1GetString(result.ToBytes());
        Assert.Equal("1.6", PdfInspector.Inspect(result.ToBytes()).CatalogVersion);
        Assert.Contains("/Version /1.6", raw, StringComparison.Ordinal);
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

    private static byte[] BuildInheritedTerminalFieldWithoutKidsPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Tx /T (section) /Kids [7 0 R] >>", "endobj",
            "7 0 obj", "<< /Parent 6 0 R /T (existing) /V (before) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
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

    private static byte[] BuildInheritedPushButtonPdf(bool includeInteractiveAppearances = true) {
        const string appearance = "BT /F1 10 Tf (Run) Tj ET";
        const string rollover = "BT /F1 10 Tf (Rollover) Tj ET";
        const string down = "BT /F1 10 Tf (Down) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [8 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Btn /Ff 65536 /T (group) /Kids [7 0 R] >>", "endobj",
            "7 0 obj", "<< /Parent 6 0 R /T (run) /Kids [8 0 R] >>", "endobj",
            "8 0 obj", "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 120 40] /P 3 0 R /MK << /CA (Run) >> /AP << /N 9 0 R " + (includeInteractiveAppearances ? "/R 10 0 R /D 11 0 R " : string.Empty) + ">> >>", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Length " + appearance.Length + " >>", "stream", appearance, "endstream", "endobj",
            "10 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Length " + rollover.Length + " >>", "stream", rollover, "endstream", "endobj",
            "11 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Length " + down.Length + " >>", "stream", down, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 12 >>", "%%EOF"
        }));
    }

    private static byte[] BuildNonzeroGenerationWidgetPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 2 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 2 R] >>", "endobj",
            "6 2 obj", "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (Ada) /Rect [20 20 120 40] /P 3 0 R >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSharedPushButtonAppearancePdf() {
        const string appearance = "BT (Shared) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R 7 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R 7 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (first) /Rect [20 20 120 40] /P 3 0 R /MK << /CA (First) >> /AP 8 0 R >>", "endobj",
            "7 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (second) /Rect [20 60 120 80] /P 3 0 R /MK << /CA (Second) >> /AP 8 0 R >>", "endobj",
            "8 0 obj", "<< /N 9 0 R >>", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 20] /Length " + appearance.Length + " >>", "stream", appearance, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }));
    }

    private static byte[] BuildCatalogVersionedOpenTypePdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Version /1.6 /Pages 2 0 R /OfficeIMOFont 4 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>", "endobj",
            "4 0 obj", "<< /Subtype /OpenType /Length 0 >>", "stream", "", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static PdfDictionary RequireNamedField(Dictionary<int, PdfIndirectObject> objects, string name) {
        return Assert.IsType<PdfDictionary>(Assert.Single(objects.Values, item =>
            item.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfStringObj>("T")?.Value == name).Value);
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

    private static byte[] BuildPageDestinationActionPdf(int firstPageObjectNumber, int secondPageObjectNumber) {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [" + firstPageObjectNumber + " 0 R " + secondPageObjectNumber + " 0 R] >>", "endobj",
            firstPageObjectNumber + " 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /AA << /O << /S /GoTo /D [" + secondPageObjectNumber + " 0 R /Fit] >> >> >>", "endobj",
            secondPageObjectNumber + " 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size " + (secondPageObjectNumber + 1) + " >>", "%%EOF"
        }));
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

    private static byte[] BuildSinglePagePdf(string version, string? catalogVersion = null) {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-" + version,
            "1 0 obj", "<< /Type /Catalog " + (catalogVersion is null ? string.Empty : "/Version /" + catalogVersion + " ") + "/Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 4 >>", "%%EOF"
        }));
    }
}
