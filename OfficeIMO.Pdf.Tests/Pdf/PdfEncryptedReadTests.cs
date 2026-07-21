using System.Security.Cryptography;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedReadTests {
    [Fact]
    public void StandardPasswordEncryptedPdf_RequiresValidPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(pdf));
        Assert.Throws<PdfPasswordRequiredException>(() => PdfTextExtractor.ExtractAllText(pdf));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Open(pdf, new PdfReadOptions { Password = "wrong" }));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "wrong" }));
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_ReadsTextWithPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");
        var options = new PdfReadOptions { Password = "open" };

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options);
        PdfDiagnosticReport diagnostics = PdfDiagnostics.Analyze(pdf, options);
        string fluentText = PdfDocument.Open(pdf, options).Read.Text();

        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.True(preflight.Probe.Security.HasEncryption);
        Assert.Contains("Secret PDF Text", text, StringComparison.Ordinal);
        Assert.Contains("Secret PDF Text", fluentText, StringComparison.Ordinal);
        Assert.True(diagnostics.ObjectGraphParsed);
        Assert.Contains(diagnostics.Findings, finding => finding.Code == "EncryptionDetected");
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_InspectUsesDecryptedSignatureSecurity() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2WithSignature("open", "owner", "Secret PDF Text");
        var options = new PdfReadOptions { Password = "open" };

        PdfDocumentInfo info = PdfInspector.Inspect(pdf, options);
        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(pdf, options);

        Assert.True(info.Security.HasSignatures);
        Assert.Equal(1, info.Security.SignatureCount);
        Assert.Equal(new[] { "Approval" }, info.Security.SignatureFieldNames);
        Assert.True(report.HasSignatures);
        Assert.Equal(1, report.SignatureCount);
        Assert.Equal("Approval", Assert.Single(report.Signatures).Signature.FieldName);
        Assert.DoesNotContain(report.Findings, finding => finding.Code == "NoSignatures");
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_PreflightReportsMissingPasswordAsReadBlocker() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.False(preflight.CanRead);
        Assert.Contains(preflight.ReadBlockers, blocker => blocker.Kind == PdfReadBlockerKind.Encryption);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.Encryption);
    }

    [Fact]
    public void StandardEmptyUserPasswordEncryptedPdf_SplitsAndExtractsWithoutExplicitPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2(string.Empty, "owner", "Empty user password text");

        byte[] extracted = PdfPageExtractor.ExtractPages(pdf, 1);
        IReadOnlyList<PdfDocument> splitPages = PdfDocument.Open(pdf).Pages.Split();
        PdfOperationResult<IReadOnlyList<PdfDocument>> trySplit = PdfDocument.Open(pdf).Pages.TrySplit();

        Assert.False(PdfInspector.Probe(extracted).HasEncryption);
        Assert.Contains("Empty user password text", PdfTextExtractor.ExtractAllText(extracted), StringComparison.Ordinal);
        PdfDocument splitPage = Assert.Single(splitPages);
        Assert.Contains("Empty user password text", splitPage.Read.Text(), StringComparison.Ordinal);
        Assert.True(trySplit.CanAttempt);
        Assert.True(trySplit.Succeeded);
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_SplitsWithPasswordAsUnencryptedOutputs() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");
        var options = new PdfReadOptions { Password = "open" };

        IReadOnlyList<byte[]> pages = PdfDocument.Open(pdf, options).Pages.Split()
            .Select(page => page.ToBytes())
            .ToArray();

        Assert.Single(pages);
        Assert.False(PdfInspector.Probe(pages[0]).HasEncryption);
        Assert.Contains("Secret PDF Text", PdfTextExtractor.ExtractAllText(pages[0]), StringComparison.Ordinal);
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_TrySplitSucceedsWhenOpenedWithPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        PdfOperationResult<IReadOnlyList<PdfDocument>> result = PdfDocument.Open(pdf, new PdfReadOptions { Password = "open" }).Pages.TrySplit();

        Assert.True(result.CanAttempt);
        Assert.True(result.Succeeded);
        PdfDocument page = Assert.Single(result.RequireValue());
        Assert.Contains("Secret PDF Text", page.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_TrySplitUsesSuppliedPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        PdfOperationResult<IReadOnlyList<PdfDocument>> result = PdfDocument.Open(pdf).Pages.TrySplit(new PdfReadOptions { Password = "open" });

        Assert.True(result.CanAttempt);
        Assert.True(result.Succeeded);
        PdfDocument page = Assert.Single(result.RequireValue());
        Assert.Contains("Secret PDF Text", page.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_ExtractWithWrongPasswordReportsPasswordError() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2WithPageLabels("open", "owner", "Secret PDF Text");

        Assert.Throws<PdfInvalidPasswordException>(() => PdfPageExtractor.ExtractPages(pdf, new PdfReadOptions { Password = "wrong" }, 1));
    }

    [Fact]
    public void StandardPasswordEncryptedPdf_ExtractsWithSupportedPageLabelsWhenPasswordProvided() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2WithPageLabels("open", "owner", "Secret PDF Text");
        var options = new PdfReadOptions { Password = "open" };

        byte[] page = PdfPageExtractor.ExtractPages(pdf, options, 1);

        Assert.False(PdfInspector.Probe(page).HasEncryption);
        Assert.Contains("Secret PDF Text", PdfTextExtractor.ExtractAllText(page), StringComparison.Ordinal);
    }

    [Fact]
    public void StandardPasswordEncryptedFormPdf_ExtractBlocksDecryptedFormMarkers() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2WithTextField("open", "owner", "Secret PDF Text");

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfPageExtractor.ExtractPages(pdf, new PdfReadOptions { Password = "open" }, 1));

        Assert.Contains("FullRewrite.Forms", exception.Plan.BlockerCodes);
    }


    [Fact]
    public void StandardPasswordEncryptedFormPdf_FillsWithValidPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2WithTextField("open", "owner", "Secret PDF Text");

        var readOptions = new PdfReadOptions { Password = "open" };
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, readOptions);
        PdfDocument filled = PdfDocument.Open(pdf, readOptions).Forms.Fill(new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.Encryption);
        Assert.False(PdfInspector.Probe(filled.ToBytes()).HasEncryption);
        Assert.Equal("Grace", Assert.Single(filled.Inspect().FormFields).Value);
    }

    private static class EncryptedPdfFixture {
        private static readonly byte[] PasswordPadding = new byte[] {
            0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
            0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
            0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
            0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
        };

        public static byte[] CreateRevision2(string userPassword, string ownerPassword, string visibleText, int encryptGeneration = 0) {
            return CreateRevision2Core(userPassword, ownerPassword, visibleText, includeSignature: false, encryptGeneration);
        }

        public static byte[] CreateRevision2WithSignature(string userPassword, string ownerPassword, string visibleText) {
            return CreateRevision2Core(userPassword, ownerPassword, visibleText, includeSignature: true, encryptGeneration: 0);
        }

        public static byte[] CreateRevision2WithPageLabels(string userPassword, string ownerPassword, string visibleText) {
            return CreateRevision2Core(userPassword, ownerPassword, visibleText, includeSignature: false, encryptGeneration: 0, includePageLabels: true);
        }

        private static byte[] CreateRevision2Core(string userPassword, string ownerPassword, string visibleText, bool includeSignature, int encryptGeneration, bool includePageLabels = false) {
            byte[] fileId = new byte[] {
                0x10, 0x45, 0xA8, 0x7C, 0x22, 0x18, 0x4E, 0xC1,
                0x91, 0x4A, 0xCF, 0x66, 0x31, 0xD2, 0x74, 0x03
            };
            const int permissions = -4;
            byte[] ownerEntry = Rc4(ComputeOwnerKey(ownerPassword), PadPassword(userPassword));
            byte[] fileKey = ComputeFileKey(userPassword, ownerEntry, permissions, fileId);
            byte[] userEntry = Rc4(fileKey, PasswordPadding);

            string content = "BT /F1 12 Tf 72 120 Td (" + EscapePdfString(visibleText) + ") Tj ET";
            byte[] encryptedContent = Rc4(ComputeObjectKey(fileKey, 5, 0), Encoding.ASCII.GetBytes(content));

            var objects = new List<byte[]>();
            string catalog = "<< /Type /Catalog /Pages 2 0 R";
            if (includeSignature) {
                catalog += " /AcroForm 7 0 R";
            }

            if (includePageLabels) {
                catalog += " /PageLabels << /Nums [0 << /S /D >>] >>";
            }

            objects.Add(Ascii(catalog + " >>"));
            objects.Add(Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
            objects.Add(Ascii(includeSignature
                ? "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R /Annots [8 0 R] >>"
                : "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"));
            objects.Add(Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));
            objects.Add(BuildStreamObject(encryptedContent));
            objects.Add(Ascii("<< /Filter /Standard /V 1 /R 2 /Length 40 /O <" + Hex(ownerEntry) + "> /U <" + Hex(userEntry) + "> /P " + permissions.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>"));
            if (includeSignature) {
                string encryptedFieldName = Hex(Rc4(ComputeObjectKey(fileKey, 8, 0), Encoding.ASCII.GetBytes("Approval")));
                string encryptedSigner = Hex(Rc4(ComputeObjectKey(fileKey, 9, 0), Encoding.ASCII.GetBytes("Alice")));
                string encryptedContents = Hex(Rc4(ComputeObjectKey(fileKey, 9, 0), new byte[] { 0x00, 0x11, 0x22 }));
                objects.Add(Ascii("<< /Fields [8 0 R] /SigFlags 3 >>"));
                objects.Add(Ascii("<< /FT /Sig /T <" + encryptedFieldName + "> /V 9 0 R /Subtype /Widget /Rect [10 10 120 40] >>"));
                objects.Add(Ascii("<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name <" + encryptedSigner + "> /ByteRange [0 8 16 30] /Contents <" + encryptedContents + "> >>"));
            }

            using var output = new MemoryStream();
            WriteAscii(output, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
            var offsets = new List<long> { 0 };
            for (int i = 0; i < objects.Count; i++) {
                int generation = i == 5 ? encryptGeneration : 0;
                offsets.Add(output.Position);
                WriteAscii(output, (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " obj\n");
                output.Write(objects[i], 0, objects[i].Length);
                WriteAscii(output, "\nendobj\n");
            }

            long xrefOffset = output.Position;
            WriteAscii(output, "xref\n0 " + (objects.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n");
            for (int i = 1; i < offsets.Count; i++) {
                int generation = i == 6 ? encryptGeneration : 0;
                WriteAscii(output, offsets[i].ToString("0000000000", System.Globalization.CultureInfo.InvariantCulture) + " " + generation.ToString("00000", System.Globalization.CultureInfo.InvariantCulture) + " n \n");
            }

            string idHex = Hex(fileId);
            WriteAscii(output, "trailer\n<< /Size " + (objects.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root 1 0 R /Encrypt 6 " + encryptGeneration.ToString(System.Globalization.CultureInfo.InvariantCulture) + " R /ID [<" + idHex + "> <" + idHex + ">] >>\nstartxref\n");
            WriteAscii(output, xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
            WriteAscii(output, "\n%%EOF\n");
            return output.ToArray();
        }

        public static byte[] CreateRevision2WithTextField(string userPassword, string ownerPassword, string visibleText) {
            byte[] fileId = new byte[] {
                0x11, 0x46, 0xA9, 0x7D, 0x23, 0x19, 0x4F, 0xC2,
                0x92, 0x4B, 0xD0, 0x67, 0x32, 0xD3, 0x75, 0x04
            };
            const int permissions = -4;
            byte[] ownerEntry = Rc4(ComputeOwnerKey(ownerPassword), PadPassword(userPassword));
            byte[] fileKey = ComputeFileKey(userPassword, ownerEntry, permissions, fileId);
            byte[] userEntry = Rc4(fileKey, PasswordPadding);

            string content = "BT /F1 12 Tf 72 120 Td (" + EscapePdfString(visibleText) + ") Tj ET";
            byte[] encryptedContent = Rc4(ComputeObjectKey(fileKey, 5, 0), Encoding.ASCII.GetBytes(content));
            string encryptedFieldName = Hex(Rc4(ComputeObjectKey(fileKey, 6, 0), Encoding.ASCII.GetBytes("Name")));
            string encryptedFieldValue = Hex(Rc4(ComputeObjectKey(fileKey, 6, 0), Encoding.ASCII.GetBytes("Ada")));

            var objects = new List<byte[]>();
            objects.Add(Ascii("<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R >>"));
            objects.Add(Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
            objects.Add(Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>"));
            objects.Add(Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));
            objects.Add(BuildStreamObject(encryptedContent));
            objects.Add(Ascii("<< /Type /Annot /Subtype /Widget /FT /Tx /T <" + encryptedFieldName + "> /V <" + encryptedFieldValue + "> /Rect [50 50 180 70] /F 4 >>"));
            objects.Add(Ascii("<< /Fields [6 0 R] >>"));
            objects.Add(Ascii("<< /Filter /Standard /V 1 /R 2 /Length 40 /O <" + Hex(ownerEntry) + "> /U <" + Hex(userEntry) + "> /P " + permissions.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>"));

            using var output = new MemoryStream();
            WriteAscii(output, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
            var offsets = new List<long> { 0 };
            for (int i = 0; i < objects.Count; i++) {
                offsets.Add(output.Position);
                WriteAscii(output, (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n");
                output.Write(objects[i], 0, objects[i].Length);
                WriteAscii(output, "\nendobj\n");
            }

            long xrefOffset = output.Position;
            WriteAscii(output, "xref\n0 9\n0000000000 65535 f \n");
            for (int i = 1; i < offsets.Count; i++) {
                WriteAscii(output, offsets[i].ToString("0000000000", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
            }

            string idHex = Hex(fileId);
            WriteAscii(output, "trailer\n<< /Size 9 /Root 1 0 R /Encrypt 8 0 R /ID [<" + idHex + "> <" + idHex + ">] >>\nstartxref\n");
            WriteAscii(output, xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
            WriteAscii(output, "\n%%EOF\n");
            return output.ToArray();
        }

        public static byte[] CreateRevision3Length40(string userPassword, string ownerPassword, string visibleText) {
            byte[] fileId = new byte[] {
                0x12, 0x47, 0xAA, 0x7E, 0x24, 0x1A, 0x50, 0xC3,
                0x93, 0x4C, 0xD1, 0x68, 0x33, 0xD4, 0x76, 0x05
            };
            const int permissions = -4;
            const int keyLengthBytes = 5;
            byte[] ownerEntry = EncryptOwnerEntryRevision3(userPassword, ownerPassword, keyLengthBytes);
            byte[] fileKey = ComputeFileKeyRevision3(userPassword, ownerEntry, permissions, fileId, keyLengthBytes);
            byte[] userEntry = ComputeUserEntryRevision3(fileKey, fileId);

            string content = "BT /F1 12 Tf 72 120 Td (" + EscapePdfString(visibleText) + ") Tj ET";
            byte[] encryptedContent = Rc4(ComputeObjectKey(fileKey, 5, 0), Encoding.ASCII.GetBytes(content));

            var objects = new List<byte[]>();
            objects.Add(Ascii("<< /Type /Catalog /Pages 2 0 R >>"));
            objects.Add(Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
            objects.Add(Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"));
            objects.Add(Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));
            objects.Add(BuildStreamObject(encryptedContent));
            objects.Add(Ascii("<< /Filter /Standard /V 2 /R 3 /Length 40 /O <" + Hex(ownerEntry) + "> /U <" + Hex(userEntry) + "> /P " + permissions.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>"));

            using var output = new MemoryStream();
            WriteAscii(output, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
            var offsets = new List<long> { 0 };
            for (int i = 0; i < objects.Count; i++) {
                offsets.Add(output.Position);
                WriteAscii(output, (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n");
                output.Write(objects[i], 0, objects[i].Length);
                WriteAscii(output, "\nendobj\n");
            }

            long xrefOffset = output.Position;
            WriteAscii(output, "xref\n0 " + (objects.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n");
            for (int i = 1; i < offsets.Count; i++) {
                WriteAscii(output, offsets[i].ToString("0000000000", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
            }

            string idHex = Hex(fileId);
            WriteAscii(output, "trailer\n<< /Size " + (objects.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root 1 0 R /Encrypt 6 0 R /ID [<" + idHex + "> <" + idHex + ">] >>\nstartxref\n");
            WriteAscii(output, xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
            WriteAscii(output, "\n%%EOF\n");
            return output.ToArray();
        }

        private static byte[] BuildStreamObject(byte[] data) {
            using var output = new MemoryStream();
            WriteAscii(output, "<< /Length " + data.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
            output.Write(data, 0, data.Length);
            WriteAscii(output, "\nendstream");
            return output.ToArray();
        }

        private static byte[] ComputeOwnerKey(string ownerPassword) {
            return Take(Md5(PadPassword(ownerPassword)), 5);
        }

        private static byte[] ComputeFileKey(string userPassword, byte[] ownerEntry, int permissions, byte[] fileId) {
            var buffer = new List<byte>();
            buffer.AddRange(PadPassword(userPassword));
            buffer.AddRange(ownerEntry);
            AppendInt32LittleEndian(buffer, permissions);
            buffer.AddRange(fileId);
            return Take(Md5(buffer.ToArray()), 5);
        }

        private static byte[] EncryptOwnerEntryRevision3(string userPassword, string ownerPassword, int keyLengthBytes) {
            byte[] ownerKey = ComputeOwnerKeyRevision3(ownerPassword, keyLengthBytes);
            byte[] current = Rc4(ownerKey, PadPassword(userPassword));
            for (int i = 1; i <= 19; i++) {
                current = Rc4(XorKey(ownerKey, i), current);
            }

            return current;
        }

        private static byte[] ComputeOwnerKeyRevision3(string ownerPassword, int keyLengthBytes) {
            byte[] digest = Md5(PadPassword(ownerPassword));
            for (int i = 0; i < 50; i++) {
                digest = Md5(Take(digest, keyLengthBytes));
            }

            return Take(digest, keyLengthBytes);
        }

        private static byte[] ComputeFileKeyRevision3(string userPassword, byte[] ownerEntry, int permissions, byte[] fileId, int keyLengthBytes) {
            var buffer = new List<byte>();
            buffer.AddRange(PadPassword(userPassword));
            buffer.AddRange(ownerEntry);
            AppendInt32LittleEndian(buffer, permissions);
            buffer.AddRange(fileId);

            byte[] current = Take(Md5(buffer.ToArray()), keyLengthBytes);
            for (int i = 0; i < 50; i++) {
                current = Md5(Take(current, keyLengthBytes));
            }

            return Take(current, keyLengthBytes);
        }

        private static byte[] ComputeUserEntryRevision3(byte[] fileKey, byte[] fileId) {
            var buffer = new List<byte>(PasswordPadding.Length + fileId.Length);
            buffer.AddRange(PasswordPadding);
            buffer.AddRange(fileId);

            byte[] value = Take(Md5(buffer.ToArray()), 16);
            value = Rc4(fileKey, value);
            for (int i = 1; i <= 19; i++) {
                value = Rc4(XorKey(fileKey, i), value);
            }

            var result = new byte[32];
            Buffer.BlockCopy(value, 0, result, 0, Math.Min(16, value.Length));
            return result;
        }

        private static byte[] ComputeObjectKey(byte[] fileKey, int objectNumber, int generation) {
            var buffer = new List<byte>(fileKey.Length + 5);
            buffer.AddRange(fileKey);
            buffer.Add((byte)(objectNumber & 0xFF));
            buffer.Add((byte)((objectNumber >> 8) & 0xFF));
            buffer.Add((byte)((objectNumber >> 16) & 0xFF));
            buffer.Add((byte)(generation & 0xFF));
            buffer.Add((byte)((generation >> 8) & 0xFF));
            return Take(Md5(buffer.ToArray()), Math.Min(fileKey.Length + 5, 16));
        }

        private static byte[] PadPassword(string password) {
            byte[] passwordBytes = Encoding.ASCII.GetBytes(password);
            var padded = new byte[32];
            int copy = Math.Min(passwordBytes.Length, 32);
            Buffer.BlockCopy(passwordBytes, 0, padded, 0, copy);
            if (copy < 32) {
                Buffer.BlockCopy(PasswordPadding, 0, padded, copy, 32 - copy);
            }

            return padded;
        }

        private static byte[] Rc4(byte[] key, byte[] data) {
            var state = new byte[256];
            for (int i = 0; i < state.Length; i++) {
                state[i] = (byte)i;
            }

            int j = 0;
            for (int i = 0; i < 256; i++) {
                j = (j + state[i] + key[i % key.Length]) & 0xFF;
                Swap(state, i, j);
            }

            var result = new byte[data.Length];
            int x = 0;
            int y = 0;
            for (int i = 0; i < data.Length; i++) {
                x = (x + 1) & 0xFF;
                y = (y + state[x]) & 0xFF;
                Swap(state, x, y);
                result[i] = (byte)(data[i] ^ state[(state[x] + state[y]) & 0xFF]);
            }

            return result;
        }

        private static byte[] XorKey(byte[] key, int value) {
            byte[] result = new byte[key.Length];
            for (int i = 0; i < key.Length; i++) {
                result[i] = (byte)(key[i] ^ value);
            }

            return result;
        }

        private static byte[] Md5(byte[] data) {
#pragma warning disable CA5351, CA1850
            using MD5 md5 = MD5.Create();
            return md5.ComputeHash(data);
#pragma warning restore CA5351, CA1850
        }

        private static void Swap(byte[] state, int left, int right) {
            byte value = state[left];
            state[left] = state[right];
            state[right] = value;
        }

        private static byte[] Take(byte[] value, int count) {
            var result = new byte[count];
            Buffer.BlockCopy(value, 0, result, 0, Math.Min(value.Length, count));
            return result;
        }

        private static void AppendInt32LittleEndian(List<byte> buffer, int value) {
            unchecked {
                buffer.Add((byte)(value & 0xFF));
                buffer.Add((byte)((value >> 8) & 0xFF));
                buffer.Add((byte)((value >> 16) & 0xFF));
                buffer.Add((byte)((value >> 24) & 0xFF));
            }
        }

        private static string EscapePdfString(string value) {
            return value.Replace("\\", "\\\\")
                .Replace("(", "\\(")
                .Replace(")", "\\)");
        }

        private static byte[] Ascii(string value) {
            return Encoding.ASCII.GetBytes(value);
        }

        private static void WriteAscii(Stream stream, string value) {
            byte[] bytes = Ascii(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string Hex(byte[] bytes) {
            var builder = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++) {
                builder.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }
    }
}
