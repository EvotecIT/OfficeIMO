using System.Security.Cryptography;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedReadTests {
    [Fact]
    public void StandardPasswordEncryptedPdf_RequiresValidPassword() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Load(pdf));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Load(pdf, new PdfReadOptions { Password = "wrong" }));
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
    public void StandardPasswordEncryptedPdf_PreflightReportsMissingPasswordAsReadBlocker() {
        byte[] pdf = EncryptedPdfFixture.CreateRevision2("open", "owner", "Secret PDF Text");

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.False(preflight.CanRead);
        Assert.Contains(preflight.ReadBlockers, blocker => blocker.Kind == PdfReadBlockerKind.Encryption);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.Encryption);
    }

    private static class EncryptedPdfFixture {
        private static readonly byte[] PasswordPadding = new byte[] {
            0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41,
            0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
            0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80,
            0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A
        };

        public static byte[] CreateRevision2(string userPassword, string ownerPassword, string visibleText) {
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
            objects.Add(Ascii("<< /Type /Catalog /Pages 2 0 R >>"));
            objects.Add(Ascii("<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
            objects.Add(Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"));
            objects.Add(Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));
            objects.Add(BuildStreamObject(encryptedContent));
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
            WriteAscii(output, "xref\n0 7\n0000000000 65535 f \n");
            for (int i = 1; i < offsets.Count; i++) {
                WriteAscii(output, offsets[i].ToString("0000000000", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
            }

            string idHex = Hex(fileId);
            WriteAscii(output, "trailer\n<< /Size 7 /Root 1 0 R /Encrypt 6 0 R /ID [<" + idHex + "> <" + idHex + ">] >>\nstartxref\n");
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
