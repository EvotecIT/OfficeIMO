using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void AuthenticatedXrefRecoveryRetainsDiscardedSignatureEvidence(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] encrypted = PdfDocument.Create(new PdfOptions().SetEncryption(new("open") {
            OwnerPassword = "owner", Algorithm = algorithm, EncryptMetadata = false
        })).Paragraph(paragraph => paragraph.Text("Readable page")).ToBytes();
        var options = new PdfLoadOptions { Password = "owner" };
        Assert.True(PdfMutationPlanner.Plan(PdfInspector.Preflight(encrypted, options), PdfMutationOperation.ChangeEncryption).CanExecute);
        string source = PdfEncoding.Latin1GetString(encrypted);
        int trailerStart = source.LastIndexOf("trailer", StringComparison.Ordinal);
        int trailerEnd = source.IndexOf("startxref", trailerStart, StringComparison.Ordinal);
        string trailer = source.Substring(trailerStart, trailerEnd - trailerStart);
        int previousXref = int.Parse(source.Substring(trailerEnd + "startxref".Length).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0]);
        trailer = System.Text.RegularExpressions.Regex.Replace(trailer, @"/Size\s+\d+", "/Size 102");
        int close = trailer.LastIndexOf(">>", StringComparison.Ordinal);
        trailer = trailer.Insert(close, " /Prev " + previousXref + " ");
        const string payload = "101 0 obj\n<< /Custom << /Type /Sig >> /Custom null >>\nendobj";
        string prefix = "\n100 0 obj\n<< /Type /Metadata /Length " + payload.Length + " >>\nstream\n";
        int outerOffset = encrypted.Length + 1;
        int recoveredOffset = encrypted.Length + prefix.Length;
        string addition = prefix + payload + "\nendstream\nendobj\n";
        int xrefOffset = encrypted.Length + addition.Length;
        addition += "xref\n100 2\n" + outerOffset.ToString("D10") + " 00000 n \n" + recoveredOffset.ToString("D10") +
            " 00000 n \n" + trailer + "startxref\n" + xrefOffset + "\n%%EOF\n";
        byte[] pdf = new byte[encrypted.Length + addition.Length];
        Buffer.BlockCopy(encrypted, 0, pdf, 0, encrypted.Length);
        Buffer.BlockCopy(Encoding.ASCII.GetBytes(addition), 0, pdf, encrypted.Length, addition.Length);
        var preflight = PdfInspector.Preflight(pdf, options);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.IncompleteObjectGraph);
        Assert.True(preflight.DocumentInfo!.HasSignatures);
        Assert.False(PdfMutationPlanner.Plan(preflight, PdfMutationOperation.ChangeEncryption).CanExecute);
        Assert.ThrowsAny<Exception>(() => PdfDocument.Load(pdf, options).Security.Decrypt("owner"));
    }

    [Theory]
    [InlineData(".5", 0.5)]
    [InlineData("-.5", -0.5)]
    [InlineData("+.5", 0.5)]
    [InlineData("5.", 5)]
    public void ValidPdfRealNumbersRetainCompleteCoverage(string number, double expected) {
        string source = Encoding.ASCII.GetString(BuildOpaqueMarkerPdf("Safe metadata", false));
        byte[] pdf = Encoding.ASCII.GetBytes(source.Replace("trailer", "6 0 obj\n<< /Custom " + number + " >>\nendobj\ntrailer"));
        var (objects, _) = PdfSyntax.ParseObjects(pdf, null, out var repair);
        Assert.Equal(expected, Assert.IsType<PdfDictionary>(objects[6].Value).Get<PdfNumber>("Custom")!.Value);
        Assert.False(repair.HasUnreadableObjects);
        Assert.True(PdfMutationPlanner.Plan(pdf, PdfMutationOperation.ChangeEncryption).CanExecute);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmptyObjectStreamDeclarationCannotSwallowSyntax(bool explicitCount) {
        const string payload = "7 0 << /Type /Sig >>";
        string source = Encoding.ASCII.GetString(BuildOpaqueMarkerPdf("Safe metadata", false));
        string obj = "6 0 obj\n<< /Type /ObjStm " + (explicitCount ? "/N 0 " : "") +
            "/First " + payload.Length + " /Length " + payload.Length + " >>\nstream\n" + payload + "\nendstream\nendobj\n";
        AssertIncompleteMutationBlocked(Encoding.ASCII.GetBytes(source.Replace("trailer", obj + "trailer")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncompleteXrefStreamCannotHideCompressedObjects(bool invalidWidths) {
        using var output = new MemoryStream();
        var offsets = new Dictionary<int, int>();
        void Write(string text) { byte[] data = Encoding.ASCII.GetBytes(text); output.Write(data, 0, data.Length); }
        void Object(int id, string body) { offsets[id] = checked((int)output.Position); Write(id + " 0 obj\n" + body + "\nendobj\n"); }
        Write("%PDF-1.7\n");
        Object(1, "<< /Type /Catalog /Pages 2 0 R >>");
        Object(2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>");
        Object(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>");
        Object(4, "<< /Length 3 >>\nstream\nq\nQ\nendstream");
        Object(5, "<< /Title (Safe metadata) >>");
        const string compressed = "8 0 << /Type /Sig /ByteRange [0 10 20 30] >>";
        Object(6, "<< /Type /ObjStm /N 1 /First 4 /Length " + compressed.Length + " >>\nstream\n" + compressed + "\nendstream");
        offsets[7] = checked((int)output.Position);
        // /Size declares object 8, but its final type-2 entry is missing from the data.
        byte[] entries = new byte[8 * 7];
        for (int id = 1; id <= 7; id++) {
            entries[id * 7] = 1;
            int offset = offsets[id];
            for (int part = 0; part < 4; part++) entries[id * 7 + 1 + part] = (byte)(offset >> (24 - part * 8));
        }
        Write("7 0 obj\n<< /Type /XRef /Root 1 0 R /Size 9 /Index [0 9] /W [" + (invalidWidths ? "2147483647 2147483647 3" : "1 4 2") +
            "] /Length " + entries.Length + " >>\nstream\n");
        output.Write(entries, 0, entries.Length);
        Write("\nendstream\nendobj\nstartxref\n" + offsets[7] + "\n%%EOF\n");
        AssertIncompleteMutationBlocked(output.ToArray());
    }

    private static void AssertIncompleteMutationBlocked(byte[] pdf) {
        var preflight = PdfInspector.Preflight(pdf);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.IncompleteObjectGraph);
        Assert.False(PdfMutationPlanner.Plan(preflight, PdfMutationOperation.ChangeEncryption).CanExecute);
        Assert.ThrowsAny<Exception>(() => PdfDocument.Load(pdf).Security.Encrypt(new("open") { OwnerPassword = "owner" }));
    }
}
