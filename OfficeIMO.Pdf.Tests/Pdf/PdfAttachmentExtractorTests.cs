using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAttachmentExtractorTests {
    [Fact]
    public void ExtractAttachments_ReadsGeneratedEmbeddedAndAssociatedFiles() {
        byte[] invoiceXml = Encoding.UTF8.GetBytes("<invoice>42</invoice>");
        byte[] sourceBytes = Encoding.UTF8.GetBytes("Source payload");

        byte[] pdf = PdfDocument.Create()
            .AttachFile("invoice.xml", invoiceXml, "application/xml", PdfAssociatedFileRelationship.Data, "Structured invoice XML")
            .AttachFile("source.txt", sourceBytes, "text/plain", PdfAssociatedFileRelationship.Source)
            .Paragraph(p => p.Text("Attachment readback proof."))
            .ToBytes();

        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(pdf);
        IReadOnlyList<PdfExtractedAttachment> documentAttachments = PdfReadDocument.Open(pdf).ExtractAttachments();
        IReadOnlyList<PdfAttachmentInfo> documentAttachmentInfos = PdfReadDocument.Open(pdf).Attachments;

        Assert.Equal(2, attachments.Count);
        Assert.Equal(2, documentAttachments.Count);
        Assert.Equal(invoiceXml.Length, documentAttachmentInfos[0].SizeBytes);
        Assert.Equal(sourceBytes.Length, documentAttachmentInfos[1].SizeBytes);

        PdfExtractedAttachment invoice = attachments[0];
        Assert.Equal("invoice.xml", invoice.Name);
        Assert.Equal("invoice.xml", invoice.FileName);
        Assert.Equal("invoice.xml", invoice.UnicodeFileName);
        Assert.Equal("Structured invoice XML", invoice.Description);
        Assert.Equal("application/xml", invoice.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, invoice.Relationship);
        Assert.Equal(invoiceXml, invoice.Bytes);
        Assert.True(invoice.FileSpecObjectNumber > 0);
        Assert.True(invoice.EmbeddedFileObjectNumber > 0);

        byte[] snapshot = invoice.Bytes;
        snapshot[0] = 0;
        Assert.Equal((byte)'<', invoice.Bytes[0]);

        PdfExtractedAttachment source = attachments[1];
        Assert.Equal("source.txt", source.FileName);
        Assert.Equal("text/plain", source.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Source, source.Relationship);
        Assert.Equal(sourceBytes, source.Bytes);
    }

    [Fact]
    public void GeneratedEmbeddedFileNameTreeSortsAttachmentNames() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .AttachFile("z.xml", Encoding.UTF8.GetBytes("<z />"), "application/xml", PdfAssociatedFileRelationship.Data)
            .AttachFile("a.xml", Encoding.UTF8.GetBytes("<a />"), "application/xml", PdfAssociatedFileRelationship.Data)
            .Paragraph(p => p.Text("Sorted attachment name tree."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Matches(@"/Names \[<612E786D6C> \d+ 0 R <7A2E786D6C> \d+ 0 R\]", content);
    }

    [Fact]
    public void ExtractAttachments_ReadsSimpleHandBuiltEmbeddedFilePdf() {
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(BuildSimpleEmbeddedFilePdf());

        PdfExtractedAttachment attachment = Assert.Single(attachments);
        Assert.Equal("note.txt", attachment.Name);
        Assert.Equal("note.txt", attachment.FileName);
        Assert.Null(attachment.UnicodeFileName);
        Assert.Null(attachment.Description);
        Assert.Null(attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Unspecified, attachment.Relationship);
        Assert.Equal(5, attachment.FileSpecObjectNumber);
        Assert.Equal(6, attachment.EmbeddedFileObjectNumber);
        Assert.Equal("note", Encoding.ASCII.GetString(attachment.Bytes));
    }

    [Fact]
    public void ExtractAttachments_DecodesFlateEmbeddedFileStreams() {
        byte[] payload = Encoding.UTF8.GetBytes("compressed attachment payload");
        byte[] pdf = BuildFlateEmbeddedFilePdf(payload);

        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(pdf));

        Assert.Equal("data.bin", attachment.FileName);
        Assert.Equal("application/octet-stream", attachment.MimeType);
        Assert.Equal("FlateDecode", attachment.Filter);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal(payload, attachment.Bytes);
    }

    [Fact]
    public void InspectAttachments_DoesNotPresentDeclaredOrEncodedSizeAsDecodedSize() {
        byte[] payload = Encoding.UTF8.GetBytes(new string('A', 4096));
        byte[] pdf = BuildFlateEmbeddedFilePdf(payload, declaredSize: 1);

        PdfAttachmentInfo info = Assert.Single(PdfReadDocument.Open(pdf).Attachments);
        PdfExtractedAttachment extracted = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(pdf));

        Assert.Equal(1, info.DeclaredSizeBytes);
        Assert.True(info.EncodedSizeBytes > 1);
        Assert.Equal(info.EncodedSizeBytes, info.SizeBytes);
        Assert.Null(info.DecodedSizeBytes);
        Assert.Equal(payload.Length, extracted.Bytes.Length);
    }

    [Fact]
    public void ExtractAttachments_ReadsCatalogAssociatedFilesWithoutEmbeddedFileNameTree() {
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(BuildAssociatedFileOnlyPdf());

        PdfExtractedAttachment attachment = Assert.Single(attachments);
        Assert.Equal("data.xml", attachment.Name);
        Assert.Equal("data.xml", attachment.FileName);
        Assert.Equal("text/xml", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal("AF", attachment.Source);
        Assert.Equal(5, attachment.FileSpecObjectNumber);
        Assert.Equal(6, attachment.EmbeddedFileObjectNumber);
        Assert.Equal("data", Encoding.ASCII.GetString(attachment.Bytes));
    }

    [Fact]
    public void ExtractAttachments_CachesSharedPayloadWithoutDroppingAliases() {
        byte[] pdf = BuildRepeatedFileAttachmentAnnotationPdf(annotationCount: 1_000);

        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(pdf);

        Assert.Equal(1_000, attachments.Count);
        Assert.Equal("alias-0.txt", attachments[0].FileName);
        Assert.Equal("alias-999.txt", attachments[999].FileName);
        Assert.All(attachments, attachment => Assert.Equal("payload", Encoding.ASCII.GetString(attachment.Bytes)));
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(pdf, "alias-999.txt"));
    }

    [Fact]
    public void PdfReadDocument_BoundsAttachmentAliasCount() {
        byte[] pdf = BuildRepeatedFileAttachmentAnnotationPdf(annotationCount: 3);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions {
                Limits = new PdfReadLimits { MaxAttachments = 2 }
            }));

        Assert.Equal(PdfReadLimitKind.Attachments, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    [Fact]
    public void PdfReadDocument_BoundsAggregateUniqueAttachmentBytes() {
        byte[] pdf = PdfDocument.Create()
            .AttachFile("first.bin", new byte[] { 1, 2, 3 })
            .AttachFile("second.bin", new byte[] { 4, 5, 6 })
            .Paragraph(p => p.Text("Aggregate attachment budget."))
            .ToBytes();

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalAttachmentBytes = 5 }
        });
        Assert.Equal(2, document.Attachments.Count);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.ExtractAttachments());

        Assert.Equal(PdfReadLimitKind.AttachmentBytes, exception.Kind);
        Assert.Equal(5, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void PdfReadDocument_BoundsMalformedPredictorFallbackAttachmentBytes() {
        var payload = new byte[64];
        for (int index = 0; index < payload.Length; index++) {
            payload[index] = (byte)index;
        }
        byte[] pdf = BuildFlateEmbeddedFilePdf(payload, malformedPredictor: true);

        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxTotalAttachmentBytes = 64 }
        });
        Assert.Single(document.Attachments);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.ExtractAttachments());

        Assert.Equal(PdfReadLimitKind.AttachmentBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void PdfReadDocument_RejectsAttachmentCountBeforeDecodingNextPayload() {
        byte[] pdf = PdfDocument.Create()
            .AttachFile("first.bin", new byte[] { 1, 2, 3 })
            .AttachFile("second.bin", new byte[] { 4, 5, 6 })
            .Paragraph(p => p.Text("Attachment count ordering."))
            .ToBytes();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, new PdfReadOptions {
                Limits = new PdfReadLimits {
                    MaxAttachments = 1,
                    MaxTotalAttachmentBytes = 3
                }
            }));

        Assert.Equal(PdfReadLimitKind.Attachments, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void ExtractAttachments_SupportsPathStreamAndDirectoryOutputs() {
        byte[] payload = Encoding.UTF8.GetBytes("directory payload");
        byte[] pdf = PdfDocument.Create()
            .AttachFile("payload.txt", payload, "text/plain", PdfAssociatedFileRelationship.Supplement)
            .Paragraph(p => p.Text("Directory extraction proof."))
            .ToBytes();

        string tempRoot = Path.Combine(Path.GetTempPath(), "officeimo-pdf-attachments-" + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(tempRoot, "input.pdf");
        string outputDirectory = Path.Combine(tempRoot, "out");

        try {
            Directory.CreateDirectory(tempRoot);
            File.WriteAllBytes(pdfPath, pdf);

            using var stream = new MemoryStream(pdf);
            Assert.Single(PdfAttachmentExtractor.ExtractAttachments(pdfPath));
            Assert.Single(PdfAttachmentExtractor.ExtractAttachments(stream));

            IReadOnlyList<string> paths = PdfAttachmentExtractor.ExtractAttachments(pdf, outputDirectory);
            string path = Assert.Single(paths);

            Assert.Equal("payload.txt", Path.GetFileName(path));
            Assert.True(File.Exists(path));
            Assert.Equal(payload, File.ReadAllBytes(path));
        } finally {
            if (Directory.Exists(tempRoot)) {
                Directory.Delete(tempRoot, recursive: true);
            }
        }
    }

    [Fact]
    public void ExtractAttachments_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfAttachmentExtractor.ExtractAttachments((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfAttachmentExtractor.ExtractAttachments((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfAttachmentExtractor.ExtractAttachments((PdfReadDocument)null!));
        Assert.Throws<ArgumentException>(() => PdfAttachmentExtractor.ExtractAttachments(" "));

        using var unreadable = new WriteOnlyStream();
        Assert.Throws<ArgumentException>(() => PdfAttachmentExtractor.ExtractAttachments(unreadable));
    }

    private static byte[] BuildSimpleEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFlateEmbeddedFilePdf(byte[] payload, bool malformedPredictor = false, int? declaredSize = null) {
        byte[] compressed = DeflateZlib(payload);
        string header = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(data.bin) 5 0 R] >> >> /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.bin) /UF (data.bin) /Desc (Payload) /AFRelationship /Data /EF << /F 6 0 R /UF 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Subtype /application#2Foctet-stream /Length " + compressed.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Filter /FlateDecode" +
                (declaredSize.HasValue ? " /Params << /Size " + declaredSize.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>" : string.Empty) +
                (malformedPredictor ? " /DecodeParms << /Predictor 12 /Columns 4 >>" : string.Empty) + " >>",
            "stream"
        });
        string footer = string.Join("\n", new[] {
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        using var output = new MemoryStream();
        WriteAscii(output, header);
        output.WriteByte((byte)'\n');
        output.Write(compressed, 0, compressed.Length);
        output.WriteByte((byte)'\n');
        WriteAscii(output, footer);
        return output.ToArray();
    }

    private static byte[] BuildAssociatedFileOnlyPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Subtype /text#2Fxml /Length 4 >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildRepeatedFileAttachmentAnnotationPdf(int annotationCount) {
        var annotationReferences = new StringBuilder();
        var annotationObjects = new StringBuilder();
        var nameEntries = new StringBuilder();
        for (int index = 0; index < annotationCount; index++) {
            int fileSpecObjectNumber = 6 + (index * 2);
            int annotationObjectNumber = fileSpecObjectNumber + 1;
            nameEntries.Append("(alias-").Append(index).Append(".txt) ").Append(fileSpecObjectNumber).Append(" 0 R ");
            annotationReferences.Append(annotationObjectNumber).Append(" 0 R ");
            annotationObjects
                .Append(fileSpecObjectNumber).Append(" 0 obj\n<< /Type /Filespec /F (alias-")
                .Append(index).Append(".txt) /EF << /F 5 0 R >> >>\nendobj\n")
                .Append(annotationObjectNumber).Append(" 0 obj\n<< /Type /Annot /Subtype /FileAttachment /FS ")
                .Append(fileSpecObjectNumber).Append(" 0 R >>\nendobj\n");
        }

        string pdf = "%PDF-1.4\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [" + nameEntries + "] >> >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [" + annotationReferences + "] >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "5 0 obj\n<< /Type /EmbeddedFile /Length 7 >>\nstream\npayload\nendstream\nendobj\n" +
            annotationObjects +
            "trailer\n<< /Root 1 0 R /Size " + (6 + (annotationCount * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\n%%EOF";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] DeflateZlib(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        output.WriteByte((byte)((adler >> 24) & 0xFF));
        output.WriteByte((byte)((adler >> 16) & 0xFF));
        output.WriteByte((byte)((adler >> 8) & 0xFF));
        output.WriteByte((byte)(adler & 0xFF));
        return output.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint ModAdler = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % ModAdler;
            b = (b + a) % ModAdler;
        }

        return (b << 16) | a;
    }

    private static void WriteAscii(Stream stream, string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private sealed class WriteOnlyStream : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => 0;
        public override long Position { get => 0; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) { }
    }
}
