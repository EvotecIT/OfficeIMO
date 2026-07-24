using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMergerPolicyTests {
    [Fact]
    public void MergeWithReport_CombinesMetadataOutlinesAndAttachmentsWithDeterministicRenames() {
        byte[] first = BuildStructuredPdf("Primary", "Primary author", null, "Primary heading", "primary payload");
        byte[] second = BuildStructuredPdf("Secondary", null, "Imported subject", "Secondary heading", "secondary payload");
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy {
                Metadata = PdfMergeStructureMode.Combine,
                Outlines = PdfMergeStructureMode.Combine,
                Attachments = PdfMergeStructureMode.Combine,
                AttachmentCollisions = PdfMergeCollisionMode.RenameIncoming
            }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        byte[] merged = result.ToBytes();
        PdfReadDocument readback = PdfReadDocument.Open(merged);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(readback);

        Assert.Equal(2, readback.Pages.Count);
        Assert.Equal("Primary", readback.Metadata.Title);
        Assert.Equal("Primary author", readback.Metadata.Author);
        Assert.Equal("Imported subject", readback.Metadata.Subject);
        Assert.Collection(readback.Outlines,
            outline => { Assert.Equal("Primary heading", outline.Title); Assert.Equal(1, outline.PageNumber); },
            outline => { Assert.Equal("Secondary heading", outline.Title); Assert.Equal(2, outline.PageNumber); });
        Assert.Collection(attachments.OrderBy(static attachment => attachment.FileName, StringComparer.Ordinal),
            attachment => { Assert.Equal("evidence.source2.txt", attachment.FileName); Assert.Equal("secondary payload", Encoding.UTF8.GetString(attachment.Bytes)); },
            attachment => { Assert.Equal("evidence.txt", attachment.FileName); Assert.Equal("primary payload", Encoding.UTF8.GetString(attachment.Bytes)); });

        Assert.Equal(2, result.Report.Sources.Count);
        Assert.All(result.Report.Sources, static source => Assert.Equal(1, source.PageCount));
        Assert.All(result.Report.Sources, static source => Assert.Equal(1, source.OutlineCount));
        Assert.All(result.Report.Sources, static source => Assert.Equal(1, source.AttachmentCount));
        PdfMergeDecision attachmentDecision = Assert.Single(result.Report.Decisions, static decision => decision.Structure == "Attachments");
        Assert.Equal(PdfMergeStructureMode.Combine, attachmentDecision.Mode);
        Assert.Equal(1, attachmentDecision.ImportedCount);
        Assert.Contains("evidence.txt -> evidence.source2.txt", Assert.Single(attachmentDecision.RenamedItems), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_DefaultAttachmentPolicyPrunesIncomingPayloadObjects() {
        byte[] first = BuildStructuredPdf("Primary", null, null, "Primary heading", "primary-only-payload");
        byte[] second = BuildStructuredPdf("Incoming", null, null, "Incoming heading", "incoming-secret-payload");

        byte[] merged = PdfMerger.Merge(first, second);

        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(merged));
        Assert.Equal("primary-only-payload", Encoding.UTF8.GetString(attachment.Bytes));
        Assert.DoesNotContain("incoming-secret-payload", PdfEncoding.Latin1GetString(merged), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_DefaultAttachmentPolicyReportsIncomingMetadataWithoutDecodingPayload() {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("Primary without attachments")).ToBytes();
        var incomingOptions = new PdfOptions().AddEmbeddedFile(
            "oversized.bin",
            Enumerable.Repeat((byte)0x41, 128).ToArray(),
            "application/octet-stream");
        byte[] second = PdfDocument.Create(incomingOptions).Paragraph(p => p.Text("Incoming attachment")).ToBytes();
        var tightAttachmentLimit = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalAttachmentBytes = 16 }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(
            new PdfMergeOptions(),
            new[] { first, second },
            new[] { new PdfReadOptions(), tightAttachmentLimit });

        Assert.Equal(1, result.Report.Sources[1].AttachmentCount);
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(result.ToBytes()));
    }

    [Fact]
    public void MergeWithReport_RejectIncomingPolicyFailsBeforeReturningArtifact() {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("First")).ToBytes();
        byte[] second = PdfDocument.Create().ViewerPreferences(preferences => preferences.HideToolbar = true).Paragraph(p => p.Text("Second")).ToBytes();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { ViewerPreferences = PdfMergeStructureMode.RejectIncoming } };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => PdfMerger.MergeWithReport(options, first, second));

        Assert.Contains("rejected incoming viewer state", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MergeWithReport_CombinesDestinationsLinksAndPageLabelsAtMergedOffsets() {
        byte[] first = BuildNavigationPdf("First", PdfPageNumberStyle.LowerRoman, "front-");
        byte[] second = BuildNavigationPdf("Second", PdfPageNumberStyle.UpperLetter, "appendix-");
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy {
                NamedDestinations = PdfMergeStructureMode.Combine,
                NamedDestinationCollisions = PdfMergeCollisionMode.RenameIncoming,
                PageLabels = PdfMergeStructureMode.Combine
            }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());

        Assert.Collection(info.NamedDestinations.OrderBy(static destination => destination.Name, StringComparer.Ordinal),
            destination => { Assert.Equal("Shared", destination.Name); Assert.Equal(1, destination.PageNumber); },
            destination => { Assert.Equal("Shared.source2", destination.Name); Assert.Equal(2, destination.PageNumber); });
        Assert.Collection(info.LinkAnnotations.OrderBy(static link => link.PageNumber),
            link => Assert.Equal("Shared", link.DestinationName),
            link => Assert.Equal("Shared.source2", link.DestinationName));
        Assert.Collection(info.PageLabels,
            label => { Assert.Equal(0, label.StartPageIndex); Assert.Equal("r", label.Style); Assert.Equal("front-", label.Prefix); },
            label => { Assert.Equal(1, label.StartPageIndex); Assert.Equal("A", label.Style); Assert.Equal("appendix-", label.Prefix); });
        PdfMergeDecision destinations = Assert.Single(result.Report.Decisions, static decision => decision.Structure == "NamedDestinations");
        Assert.Contains("Shared -> Shared.source2", Assert.Single(destinations.RenamedItems), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_DefaultPolicyRemovesIncomingLinksToDroppedNamedDestinations() {
        byte[] first = BuildNavigationPdf("First", PdfPageNumberStyle.Arabic, "first-");
        byte[] second = BuildNavigationPdf("Second", PdfPageNumberStyle.Arabic, "second-");

        PdfDocumentInfo info = PdfInspector.Inspect(PdfMerger.Merge(first, second));

        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);
        Assert.Equal("Shared", destination.Name);
        Assert.Equal(1, destination.PageNumber);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal("Shared", link.DestinationName);
    }

    [Fact]
    public void Merge_KeepFirstDestinationCollisionRemovesSilentlyRetargetedIncomingLinks() {
        byte[] first = BuildNavigationPdf("First", PdfPageNumberStyle.Arabic, "first-");
        byte[] second = BuildNavigationPdf("Second", PdfPageNumberStyle.Arabic, "second-");
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy {
                NamedDestinations = PdfMergeStructureMode.Combine,
                NamedDestinationCollisions = PdfMergeCollisionMode.KeepFirst
            }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());

        Assert.Single(info.NamedDestinations);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, link.PageNumber);
        PdfMergeDecision decision = Assert.Single(result.Report.Decisions, static item => item.Structure == "NamedDestinations");
        Assert.Equal(1, decision.DroppedCount);
    }

    [Fact]
    public void MergeWithReport_CombinesSimpleAcroFormsAndRenamesIncomingFields() {
        byte[] first = PdfDocument.Create().TextField("Shared", value: "first").ToBytes();
        byte[] second = PdfDocument.Create().TextField("Shared", value: "second").ToBytes();
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy {
                Forms = PdfMergeStructureMode.Combine,
                FormFieldCollisions = PdfMergeCollisionMode.RenameIncoming
            }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        PdfReadDocument readback = PdfReadDocument.Open(result.ToBytes());

        Assert.Collection(readback.FormFields.OrderBy(static field => field.Name, StringComparer.Ordinal),
            field => { Assert.Equal("Shared", field.Name); Assert.Equal("first", field.Value); Assert.Equal(new[] { 1 }, field.PageNumbers); },
            field => { Assert.Equal("Shared.source2", field.Name); Assert.Equal("second", field.Value); Assert.Equal(new[] { 2 }, field.PageNumbers); });
        PdfMergeDecision forms = Assert.Single(result.Report.Decisions, static decision => decision.Structure == "Forms");
        Assert.Equal(1, forms.ImportedCount);
        Assert.Contains("Shared -> Shared.source2", Assert.Single(forms.RenamedItems), StringComparison.Ordinal);
    }

    [Fact]
    public void MergeWithReport_CombinesButtonChoiceAndRadioFields() {
        byte[] first = BuildCommonFormPdf("first");
        byte[] second = BuildCommonFormPdf("second");
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy { Forms = PdfMergeStructureMode.Combine }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        IReadOnlyList<PdfFormField> fields = PdfReadDocument.Open(result.ToBytes()).FormFields;

        Assert.Equal(6, fields.Count);
        Assert.Contains(fields, static field => field.Name == "Approved" && field.Kind == PdfFormFieldKind.Button);
        Assert.Contains(fields, static field => field.Name == "Approved.source2" && field.Kind == PdfFormFieldKind.Button);
        Assert.Contains(fields, static field => field.Name == "Region" && field.Kind == PdfFormFieldKind.Choice);
        Assert.Contains(fields, static field => field.Name == "Region.source2" && field.Kind == PdfFormFieldKind.Choice);
        Assert.Contains(fields, static field => field.Name == "Plan" && field.IsRadioButton);
        Assert.Contains(fields, static field => field.Name == "Plan.source2" && field.IsRadioButton);
    }

    [Fact]
    public void Merge_DefaultPolicyKeepsOnlyPrimaryAcroFormWidgets() {
        byte[] first = PdfDocument.Create().TextField("Primary", value: "first").ToBytes();
        byte[] second = PdfDocument.Create().TextField("Incoming", value: "second").ToBytes();

        byte[] merged = PdfMerger.Merge(first, second);
        PdfReadDocument readback = PdfReadDocument.Open(merged);

        PdfFormField field = Assert.Single(readback.FormFields);
        Assert.Equal("Primary", field.Name);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.Empty(PdfInspector.Inspect(merged).GetFormWidgets(2));
    }

    [Fact]
    public void Merge_DefaultPolicyPreservesPrimaryFieldsWithoutPageWidgets() {
        byte[] primary = BuildHiddenFormFieldPdf();
        byte[] incoming = PdfDocument.Create().TextField("Incoming", value: "second").ToBytes();

        PdfReadDocument readback = PdfReadDocument.Open(PdfMerger.Merge(primary, incoming));

        PdfFormField field = Assert.Single(readback.FormFields);
        Assert.Equal("Hidden", field.Name);
        Assert.Equal("secret", field.Value);
        Assert.Empty(field.PageNumbers);
    }

    [Fact]
    public void Merge_PreservesSiblingFieldsUnderHierarchicalAcroFormRoot() {
        byte[] source = BuildHierarchicalSiblingFormPdf();

        IReadOnlyList<PdfFormField> fields = PdfReadDocument.Open(PdfMerger.Merge(source)).FormFields;

        Assert.Collection(fields.OrderBy(static field => field.Name, StringComparer.Ordinal),
            field => { Assert.Equal("Parent.First", field.Name); Assert.Equal("one", field.Value); },
            field => { Assert.Equal("Parent.Second", field.Name); Assert.Equal("two", field.Value); });
    }

    [Fact]
    public void Merge_MaterializesSourceAcroFormDefaultsOnImportedFields() {
        byte[] primary = PdfDocument.Create().TextField("Primary", value: "one").ToBytes();
        byte[] incoming = BuildFormWithAcroFormDefaultsPdf();
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy { Forms = PdfMergeStructureMode.Combine }
        };

        byte[] merged = PdfMerger.MergeWithReport(options, primary, incoming).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(merged);
        PdfDictionary field = objects.Values.Select(static item => item.Value).OfType<PdfDictionary>()
            .Single(dictionary => dictionary.Get<PdfStringObj>("T")?.Value == "Incoming");

        Assert.Equal("/F1 9 Tf 0 g", field.Get<PdfStringObj>("DA")?.Value);
        Assert.Equal(2, field.Get<PdfNumber>("Q")?.Value);
        PdfDictionary resources = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, field.Items["DR"]));
        PdfDictionary fonts = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, resources.Items["Font"]));
        PdfDictionary font = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fonts.Items["F1"]));
        Assert.Equal("Helvetica", font.Get<PdfName>("BaseFont")?.Name);
    }

    [Fact]
    public void Merge_OutlineCombinePromotesChildrenOfDestinationlessParents() {
        byte[] source = BuildDestinationlessOutlineParentPdf();
        var options = new PdfMergeOptions {
            Policy = new PdfMergePolicy { Outlines = PdfMergeStructureMode.Combine }
        };

        PdfReadDocument readback = PdfReadDocument.Open(PdfMerger.MergeWithReport(options, source).ToBytes());

        PdfOutlineItem outline = Assert.Single(readback.Outlines);
        Assert.Equal("Child", outline.Title);
        Assert.Equal(1, outline.PageNumber);
    }

    [Fact]
    public void MergeWithReport_CombinesViewerPreferencesAndRetargetsIncomingOpenAction() {
        byte[] first = PdfDocument.Create()
            .CatalogView(PdfCatalogPageMode.UseThumbs, PdfCatalogPageLayout.OneColumn)
            .ViewerPreferences(preferences => preferences.HideToolbar = true)
            .Paragraph(p => p.Text("First"))
            .ToBytes();
        byte[] second = PdfDocument.Create()
            .ViewerPreferences(preferences => preferences.HideMenubar = true)
            .OpenAction(1, destinationTop: 500)
            .Paragraph(p => p.Text("Second"))
            .ToBytes();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { ViewerPreferences = PdfMergeStructureMode.Combine } };

        PdfReadDocument readback = PdfReadDocument.Open(PdfMerger.MergeWithReport(options, first, second).ToBytes());

        Assert.Equal("UseThumbs", readback.CatalogPageMode);
        Assert.Equal("OneColumn", readback.CatalogPageLayout);
        Assert.True(readback.ViewerPreferences!.GetBoolean("HideToolbar"));
        Assert.True(readback.ViewerPreferences.GetBoolean("HideMenubar"));
        Assert.Equal(2, readback.OpenAction!.PageNumber);
        Assert.Equal(500, readback.OpenAction.DestinationTop);
    }

    [Fact]
    public void MergeWithReport_PreservesIncomingOpenActionZoom() {
        byte[] first = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("First")).ToBytes();
        byte[] second = BuildRawPdf(
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /XYZ 12 100 1.5] >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream");
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { ViewerPreferences = PdfMergeStructureMode.Combine } };

        PdfDocumentOpenAction openAction = Assert.IsType<PdfDocumentOpenAction>(
            PdfInspector.Inspect(PdfMerger.MergeWithReport(options, first, second).ToBytes()).OpenAction);

        Assert.Equal(2, openAction.PageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, openAction.DestinationMode);
        Assert.Equal(12d, openAction.DestinationLeft);
        Assert.Equal(100d, openAction.DestinationTop);
        Assert.Equal(1.5d, openAction.DestinationZoom);
    }

    [Fact]
    public void MergeWithReport_CombinesCompatibleCatalogStateAndOutputIntents() {
        byte[] first = PdfDocument.Create()
            .Language("en-US")
            .CatalogUriBase("https://primary.example/")
            .SrgbOutputIntent()
            .Paragraph(p => p.Text("First"))
            .ToBytes();
        byte[] second = PdfDocument.Create()
            .Language("de-DE")
            .CatalogUriBase("https://incoming.example/")
            .SrgbOutputIntent()
            .Paragraph(p => p.Text("Second"))
            .ToBytes();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { CatalogState = PdfMergeStructureMode.Combine } };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, first, second);
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());

        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.Equal(2, info.OutputIntentCount);
        PdfMergeDecision catalog = Assert.Single(result.Report.Decisions, static decision => decision.Structure == "CatalogState");
        Assert.Contains("output-intent", catalog.Action, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeWithReport_RejectsOptionalContentCombineThatCouldExposeHiddenLayers() {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("First")).ToBytes();
        byte[] second = PdfOptionalContentSupport.BuildOptionalContentMetadataPdf();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { CatalogState = PdfMergeStructureMode.Combine } };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfMerger.MergeWithReport(options, first, second));

        Assert.Contains("intentionally hid", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(PdfMergeStructureMode.KeepPrimary)]
    [InlineData(PdfMergeStructureMode.Drop)]
    public void MergeWithReport_RejectsPoliciesThatDiscardIncomingHiddenLayerState(PdfMergeStructureMode mode) {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("First")).ToBytes();
        byte[] second = PdfDocument.Create()
            .Layer("Hidden evidence", layer => layer.Paragraph(p => p.Text("Must stay hidden")), new PdfLayerOptions {
                InitiallyVisible = false
            })
            .ToBytes();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { CatalogState = mode } };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            PdfMerger.MergeWithReport(options, first, second));

        Assert.Contains("hidden", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MergeWithReport_ReportsPageNormalizationChoice() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Source")).ToBytes();
        var options = new PdfMergeOptions {
            ResizePages = new PdfPageResizeOptions(PageSizes.A4) { Mode = PdfPageResizeMode.Fit }
        };

        PdfMergeResult result = PdfMerger.MergeWithReport(options, source);

        Assert.Single(result.Report.Decisions, static decision => decision.Structure == "PageSizeNormalization");
        Assert.Equal(595, Math.Round(PdfInspector.Inspect(result.ToBytes()).Pages[0].Width));
    }

    private static byte[] BuildStructuredPdf(string title, string? author, string? subject, string heading, string attachmentPayload) {
        var options = new PdfOptions { CreateOutlineFromHeadings = true };
        options.AddEmbeddedFile("evidence.txt", Encoding.UTF8.GetBytes(attachmentPayload), "text/plain", PdfAssociatedFileRelationship.Data, heading);
        return PdfDocument.Create(options)
            .Meta(title: title, author: author, subject: subject)
            .H1(heading)
            .Paragraph(p => p.Text(heading + " body"))
            .ToBytes();
    }

    private static byte[] BuildNavigationPdf(string text, PdfPageNumberStyle labelStyle, string labelPrefix) {
        var options = new PdfOptions();
        options.AddPageLabelRange(1, labelStyle, 1, labelPrefix);
        return PdfDocument.Create(options)
            .Paragraph(p => p.LinkToBookmark("Jump", "Shared"))
            .Bookmark("Shared")
            .Paragraph(p => p.Text(text))
            .ToBytes();
    }

    private static byte[] BuildCommonFormPdf(string marker) {
        return PdfDocument.Create()
            .CheckBox("Approved", isChecked: marker == "first")
            .ChoiceField("Region", new[] { "EU", "US" }, value: marker == "first" ? "EU" : "US")
            .RadioButtonGroup("Plan", new[] { "Basic", "Pro" }, value: marker == "first" ? "Basic" : "Pro")
            .ToBytes();
    }

    private static byte[] BuildHiddenFormFieldPdf() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Fields [6 0 R] >>",
        "<< /FT /Tx /T (Hidden) /V (secret) >>");

    private static byte[] BuildHierarchicalSiblingFormPdf() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R /Annots [9 0 R 10 0 R] >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Fields [6 0 R] >>",
        "<< /T (Parent) /Kids [7 0 R 8 0 R] >>",
        "<< /FT /Tx /T (First) /V (one) /Parent 6 0 R /Kids [9 0 R] >>",
        "<< /FT /Tx /T (Second) /V (two) /Parent 6 0 R /Kids [10 0 R] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [10 70 100 90] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [10 30 100 50] >>");

    private static byte[] BuildFormWithAcroFormDefaultsPdf() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Fields [6 0 R] /DA (/F1 9 Tf 0 g) /Q 2 /DR << /Font << /F1 7 0 R >> >> >>",
        "<< /FT /Tx /T (Incoming) /V (two) >>",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

    private static byte[] BuildDestinationlessOutlineParentPdf() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R >>",
        "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
        "<< /Title (Group) /Parent 5 0 R /First 7 0 R /Last 7 0 R /Count 1 >>",
        "<< /Title (Child) /Parent 6 0 R /Dest [3 0 R /Fit] >>");

    private static byte[] BuildRawPdf(params string[] objectBodies) {
        var builder = new StringBuilder("%PDF-1.7\n");
        for (int objectIndex = 0; objectIndex < objectBodies.Length; objectIndex++) {
            builder.Append(objectIndex + 1).Append(" 0 obj\n")
                .Append(objectBodies[objectIndex]).Append("\nendobj\n");
        }

        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objectBodies.Length + 1)
            .Append(" >>\nstartxref\n0\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
