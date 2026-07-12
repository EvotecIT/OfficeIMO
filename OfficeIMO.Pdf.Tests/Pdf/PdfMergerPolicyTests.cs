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
        PdfReadDocument readback = PdfReadDocument.Load(merged);
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
        PdfReadDocument readback = PdfReadDocument.Load(result.ToBytes());

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
        IReadOnlyList<PdfFormField> fields = PdfReadDocument.Load(result.ToBytes()).FormFields;

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
        PdfReadDocument readback = PdfReadDocument.Load(merged);

        PdfFormField field = Assert.Single(readback.FormFields);
        Assert.Equal("Primary", field.Name);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.Empty(PdfInspector.Inspect(merged).GetFormWidgets(2));
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

        PdfReadDocument readback = PdfReadDocument.Load(PdfMerger.MergeWithReport(options, first, second).ToBytes());

        Assert.Equal("UseThumbs", readback.CatalogPageMode);
        Assert.Equal("OneColumn", readback.CatalogPageLayout);
        Assert.True(readback.ViewerPreferences!.GetBoolean("HideToolbar"));
        Assert.True(readback.ViewerPreferences.GetBoolean("HideMenubar"));
        Assert.Equal(2, readback.OpenAction!.PageNumber);
        Assert.Equal(500, readback.OpenAction.DestinationTop);
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
    public void MergeWithReport_RebuildsIncomingOptionalContentAsVisibleLayers() {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("First")).ToBytes();
        byte[] second = PdfOptionalContentSupport.BuildOptionalContentMetadataPdf();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { CatalogState = PdfMergeStructureMode.Combine } };

        PdfDocumentInfo info = PdfInspector.Inspect(PdfMerger.MergeWithReport(options, first, second).ToBytes());

        Assert.Equal(new[] { "Print layer", "Hidden layer" }, info.OptionalContentGroupNames);
        Assert.All(info.OptionalContentGroups, static group => Assert.True(group.IsInitiallyVisible));
        Assert.Equal("Merged layers", info.OptionalContent!.DefaultConfigurationName);
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
}
