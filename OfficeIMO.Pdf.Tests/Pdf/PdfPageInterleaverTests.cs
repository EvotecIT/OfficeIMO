using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageInterleaverTests {
    [Fact]
    public void Interleave_AlternatesPagesAndAppendsRemainderWithProvenance() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two", "A three");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one", "B two");

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(first, "A"), new PdfInterleaveSource(second, "B") });

        Assert.Equal(5, result.Pages.Count);
        Assert.Equal(new[] { "A", "B", "A", "B", "A" }, result.Pages.Select(static page => page.SourceName));
        Assert.Equal(new[] { 1, 1, 2, 2, 3 }, result.Pages.Select(static page => page.SourcePageNumber));
        Assert.Equal(
            new[] { "Aone", "Bone", "Atwo", "Btwo", "Athree" },
            PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
        Assert.Equal(5, PdfInspector.Inspect(result.ToBytes()).PageCount);
    }

    [Fact]
    public void Interleave_HonorsReverseSelectionAndRejectsUnevenInputsWhenRequested() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one", "B two");
        var reversed = new PdfInterleaveSource(second, "B") { Reverse = true };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(first, "A"), reversed },
            new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject });

        Assert.Equal(
            new[] { "Aone", "Btwo", "Atwo", "Bone" },
            PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
        Assert.Throws<InvalidOperationException>(() => PdfPageInterleaver.Interleave(
            new[] {
                new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("one")),
                new PdfInterleaveSource(first)
            },
            new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject }));
    }

    [Fact]
    public void Interleave_ReportsOnlySelectedPagesAsImported() {
        var selected = new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("A one", "A two", "A three")) {
            Pages = PdfPageSelector.Parse("2")
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { selected, new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("B one")) });

        Assert.Equal(1, result.MergeReport.Sources[0].PageCount);
        Assert.Equal(1, result.MergeReport.Sources[1].PageCount);
        Assert.Equal(new[] { "Atwo", "Bone" }, PdfProductionWorkflowTestSupport.ReadPageTexts(result.ToBytes()));
    }

    [Fact]
    public void Interleave_PrunesPrimaryFormFieldsFromExcludedPages() {
        byte[] primary = PdfDocument.Create()
            .TextField("Kept", value: "one")
            .PageBreak()
            .TextField("Excluded", value: "two")
            .ToBytes();
        var selectedPrimary = new PdfInterleaveSource(primary) { Pages = PdfPageSelector.Parse("1") };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { selectedPrimary, new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("Incoming")) });

        PdfFormField field = Assert.Single(PdfReadDocument.Open(result.ToBytes()).FormFields);
        Assert.Equal("Kept", field.Name);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.Equal(1, result.MergeReport.Sources[0].FormFieldCount);
    }

    [Fact]
    public void Interleave_PrunesExcludedWidgetsFromAFieldSharedAcrossPages() {
        var selectedPrimary = new PdfInterleaveSource(BuildTwoPageSharedFieldPdf()) {
            Pages = PdfPageSelector.Parse("1")
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { selectedPrimary, new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("Incoming")) });

        PdfFormField field = Assert.Single(PdfReadDocument.Open(result.ToBytes()).FormFields);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal("Shared", field.Name);
        Assert.Equal(1, widget.PageNumber);
    }

    [Fact]
    public void Interleave_UsesOutputPageOwnershipWhenDroppingIncomingNamedLinks() {
        byte[] primary = PdfDocument.Create()
            .Bookmark("PrimaryDestination")
            .Paragraph(paragraph => paragraph.LinkToBookmark("Primary link one", "PrimaryDestination"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.LinkToBookmark("Primary link two", "PrimaryDestination"))
            .ToBytes();
        byte[] incoming = PdfDocument.Create()
            .Bookmark("IncomingDestination")
            .Paragraph(paragraph => paragraph.LinkToBookmark("Incoming link", "IncomingDestination"))
            .ToBytes();

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(primary), new PdfInterleaveSource(incoming) });
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());

        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);
        Assert.Equal("PrimaryDestination", destination.Name);
        Assert.Collection(
            info.LinkAnnotations.OrderBy(static link => link.PageNumber),
            link => { Assert.Equal(1, link.PageNumber); Assert.Equal("PrimaryDestination", link.DestinationName); },
            link => { Assert.Equal(3, link.PageNumber); Assert.Equal("PrimaryDestination", link.DestinationName); });
    }

    [Fact]
    public void Interleave_RejectsXfaSourcesBeforeComposition() {
        var xfaSource = new PdfInterleaveSource(BuildRawPdf(
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Fields [] /XFA (unsupported-packet) >>"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfPageInterleaver.Interleave(
            new[] { xfaSource, new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("Incoming")) }));

        Assert.Contains("XFA", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Interleave_ComposesSourceBudgetsForTheOwnedOutput() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one");
        int sourceObjectLimit = Math.Max(PdfSyntax.ParseObjects(first).Map.Count, PdfSyntax.ParseObjects(second).Map.Count);
        var firstSource = new PdfInterleaveSource(first) {
            ReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 1, MaxIndirectObjects = sourceObjectLimit } }
        };
        var secondSource = new PdfInterleaveSource(second) {
            ReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 1, MaxIndirectObjects = sourceObjectLimit } }
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(new[] { firstSource, secondSource });

        Assert.Equal(2, result.ToDocument().Reader.Pages().Count);
        Assert.Equal(2, result.MergeReport.OutputPageCount);
    }

    [Fact]
    public void Interleave_CarriesTightSourceLimitsThroughResizePreparation() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("A one");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one");
        int firstObjectCount = PdfReadDocument.Open(first).RawStructure().TotalObjectCount;
        int secondObjectCount = PdfReadDocument.Open(second).RawStructure().TotalObjectCount;
        var firstSource = new PdfInterleaveSource(first) {
            ReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = firstObjectCount } }
        };
        var secondSource = new PdfInterleaveSource(second) {
            ReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = secondObjectCount } }
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { firstSource, secondSource },
            new PdfInterleaveOptions {
                MergeOptions = new PdfMergeOptions {
                    ResizePages = new PdfPageResizeOptions(new PageSize(420, 595))
                }
            });

        Assert.Equal(2, result.ToDocument().Reader.Pages().Count);
    }

    [Fact]
    public void Interleave_ScopesIncomingPagePoliciesAndLinksToSelectedPages() {
        byte[] incoming = BuildRawPdf(
            "<< /Type /Catalog /Pages 2 0 R /Outlines 8 0 R /PageMode /UseOutlines /PageLabels 10 0 R /OpenAction [5 0 R /Fit] >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R /Annots [7 0 R] >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 6 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] /Dest [5 0 R /Fit] /Contents (excluded-link-owner) >>",
            "<< /Type /Outlines /First 9 0 R /Last 9 0 R /Count 1 >>",
            "<< /Title (Excluded outline) /Parent 8 0 R /Dest [5 0 R /Fit] >>",
            "<< /Nums [1 << /S /D /St 1 >>] >>");
        var selectedIncoming = new PdfInterleaveSource(incoming) { Pages = PdfPageSelector.Parse("1") };
        var options = new PdfInterleaveOptions {
            MergeOptions = new PdfMergeOptions {
                Policy = new PdfMergePolicy {
                    Outlines = PdfMergeStructureMode.RejectIncoming,
                    PageLabels = PdfMergeStructureMode.RejectIncoming,
                    ViewerPreferences = PdfMergeStructureMode.RejectIncoming
                }
            }
        };

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] {
                new PdfInterleaveSource(PdfProductionWorkflowTestSupport.CreatePdf("Primary")),
                selectedIncoming
            },
            options);

        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());
        Assert.Empty(info.LinkAnnotations);
        Assert.Equal(0, result.MergeReport.Sources[1].OutlineCount);
        Assert.Equal(0, result.MergeReport.Sources[1].PageLabelCount);
        Assert.DoesNotContain("excluded-link-owner", System.Text.Encoding.ASCII.GetString(result.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void Interleave_MapsCombinedOpenActionsThroughRoundRobinOutputOrder() {
        byte[] first = BuildRawPdf(
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [5 0 R /Fit] >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 6 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("B one", "B two");

        PdfInterleaveResult result = PdfPageInterleaver.Interleave(
            new[] { new PdfInterleaveSource(first), new PdfInterleaveSource(second) },
            new PdfInterleaveOptions {
                MergeOptions = new PdfMergeOptions {
                    Policy = new PdfMergePolicy { ViewerPreferences = PdfMergeStructureMode.Combine }
                }
            });

        Assert.Equal(3, PdfInspector.Inspect(result.ToBytes()).OpenAction!.PageNumber);
    }


    private static byte[] BuildTwoPageSharedFieldPdf() => BuildRawPdf(
        "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R >>",
        "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R /Annots [9 0 R] >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 6 0 R /Annots [10 0 R] >>",
        "<< /Length 0 >>\nstream\n\nendstream",
        "<< /Fields [8 0 R] >>",
        "<< /FT /Tx /T (Shared) /V (value) /Kids [9 0 R 10 0 R] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 8 0 R /P 3 0 R /Rect [10 70 100 90] >>",
        "<< /Type /Annot /Subtype /Widget /Parent 8 0 R /P 5 0 R /Rect [10 30 100 50] >>");

    private static byte[] BuildRawPdf(params string[] objectBodies) {
        var builder = new System.Text.StringBuilder("%PDF-1.7\n");
        for (int objectIndex = 0; objectIndex < objectBodies.Length; objectIndex++) {
            builder.Append(objectIndex + 1).Append(" 0 obj\n")
                .Append(objectBodies[objectIndex]).Append("\nendobj\n");
        }
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objectBodies.Length + 1)
            .Append(" >>\nstartxref\n0\n%%EOF\n");
        return System.Text.Encoding.ASCII.GetBytes(builder.ToString());
    }
}
