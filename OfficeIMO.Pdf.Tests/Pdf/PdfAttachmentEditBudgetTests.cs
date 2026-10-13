using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAttachmentEditBudgetTests {
    [Fact]
    public void AddRejectsTheTargetBudgetBeforeChangingTheSessionAndAllowsRecovery() {
        byte[] source = CreateSource();
        var document = PdfDocument.Load(source, ReadOptions(4));

        PdfAttachmentEditResult result = document.Attachments.Edit(session => {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                session.Add(new PdfEmbeddedFile("too-large.bin", new byte[5])));
            Assert.Equal(PdfReadLimitKind.AttachmentBytes, error.Kind);
            Assert.Equal(4, error.Limit);
            Assert.Equal(5, error.Actual);
            Assert.Empty(session.Attachments);
            session.Add(new PdfEmbeddedFile("fits.bin", new byte[] { 1, 2, 3, 4 }));
        });

        Assert.True(result.IsValid);
        var reopened = PdfDocument.Load(result.ToBytes(), ReadOptions(4));
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, Assert.Single(reopened.Attachments.Extract()).Bytes);
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void ReplaceAdmitsOnlyTheNetPayloadAndLeavesNamesAndDataIntactOnRejection() {
        byte[] source = CreateSource();
        var initial = PdfDocument.Load(source);
        source = initial.Attachments.Edit(session => session
            .Add(new PdfEmbeddedFile("first.bin", new byte[] { 1, 2, 3 }))
            .Add(new PdfEmbeddedFile("second.bin", new byte[] { 4, 5 }))).ToBytes();
        var document = PdfDocument.Load(source, ReadOptions(5));

        PdfAttachmentEditResult result = document.Attachments.Edit(session => {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                session.Replace("first.bin", new PdfEmbeddedFile("rejected.bin", new byte[4])));
            Assert.Equal(6, error.Actual);
            Assert.Equal(new[] { "first.bin", "second.bin" }, session.Attachments.Select(file => file.FileName));
            Assert.Equal(new byte[] { 1, 2, 3 }, session.Attachments[0].Data);
            session.Replace("first.bin", new PdfEmbeddedFile("accepted.bin", new byte[] { 6, 7, 8 }));
        });

        Assert.True(result.IsValid);
        var reopened = PdfDocument.Load(result.ToBytes(), ReadOptions(5));
        Assert.Equal(new[] { "accepted.bin", "second.bin" }, reopened.Attachments.Extract().Select(file => file.FileName));
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void RemoveReleasesTheTargetBudgetForLaterAdds() {
        var initial = PdfDocument.Load(CreateSource());
        byte[] source = initial.Attachments.Add(new PdfEmbeddedFile("old.bin", new byte[4])).ToBytes();
        var document = PdfDocument.Load(source, ReadOptions(4));

        PdfAttachmentEditResult result = document.Attachments.Edit(session => session
            .Remove("old.bin")
            .Add(new PdfEmbeddedFile("new.bin", new byte[] { 4, 3, 2, 1 })));

        var reopened = PdfDocument.Load(result.ToBytes(), ReadOptions(4));
        PdfExtractedAttachment attachment = Assert.Single(reopened.Attachments.Extract());
        Assert.Equal("new.bin", attachment.FileName);
        Assert.Equal(new byte[] { 4, 3, 2, 1 }, attachment.Bytes);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ExplicitEditBudgetAlsoBoundsNewPayloads() {
        byte[] source = CreateSource();
        PdfAttachmentEditResult result = PdfAttachmentEditor.Edit(source, session => {
            PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
                session.Add(new PdfEmbeddedFile("too-large.bin", new byte[5])));
            Assert.Equal(4, error.Limit);
            session.Add(new PdfEmbeddedFile("fits.bin", new byte[4]));
        }, ReadOptions(16), maxDecodedAttachmentBytes: 4);

        Assert.True(result.IsValid);
    }

    private static PdfLoadOptions ReadOptions(long maximumAttachmentBytes) =>
        new() { Limits = new PdfReadLimits { MaxTotalAttachmentBytes = maximumAttachmentBytes } };

    private static byte[] CreateSource() => PdfDocument.Create(pdf =>
        pdf.Content(content => content.Paragraph(paragraph => paragraph.Text("Attachment budget proof")))).ToBytes();
}
