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
    public void MergeWithReport_RejectsPolicyModesThatAreNotYetEnforced() {
        byte[] first = PdfDocument.Create().Paragraph(p => p.Text("First")).ToBytes();
        byte[] second = PdfDocument.Create().Paragraph(p => p.Text("Second")).ToBytes();
        var options = new PdfMergeOptions { Policy = new PdfMergePolicy { NamedDestinations = PdfMergeStructureMode.Combine } };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfMerger.MergeWithReport(options, first, second));

        Assert.Contains("NamedDestinations=Combine", exception.Message, StringComparison.Ordinal);
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
}
