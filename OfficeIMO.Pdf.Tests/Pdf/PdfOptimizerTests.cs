using System.Text;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOptimizerTests {
    [Fact]
    public void Optimize_CompressesUnfilteredStreamsAndPreservesText() {
        byte[] source = BuildPdfWithUncompressedTextStream("BT\n/F1 12 Tf\n72 720 Td\n(" + new string('A', 4096) + ") Tj\nET\n");

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source);

        Assert.True(result.Applied);
        Assert.True(result.OptimizedLengthBytes < result.OriginalLengthBytes);
        Assert.NotNull(result.ReportAfter);
        Assert.True(result.CandidateLengthBytes <= result.OriginalLengthBytes);
        Assert.True(result.CandidateSavedBytes > 0);
        Assert.True(result.SkippedActionCount >= 0);
        PdfOptimizationAction action = Assert.Single(result.Actions);
        Assert.Equal("CompressStream", action.Kind);
        Assert.Equal(5, action.ObjectNumber);
        Assert.Contains("/Filter /FlateDecode", PdfEncoding.Latin1GetString(result.Bytes), StringComparison.Ordinal);
        byte[] compressedStream = ExtractFirstStreamData(result.Bytes);
        Assert.True(compressedStream.Length > 2);
        Assert.Equal(0x78, compressedStream[0]);
        Assert.Equal(0, ((compressedStream[0] << 8) + compressedStream[1]) % 31);
        Assert.Contains(new string('A', 64), PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Optimize_ReturnsOriginalWhenCandidateIsNotSmaller() {
        byte[] source = BuildPdfWithUncompressedTextStream("BT\n/F1 12 Tf\n72 720 Td\n(Tiny) Tj\nET\n");

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source);

        Assert.False(result.Applied);
        Assert.True(result.ReturnedOriginal);
        Assert.True(result.CandidateLengthBytes >= result.OriginalLengthBytes);
        Assert.Equal(0, result.CandidateSavedBytes);
        Assert.Contains(result.SkippedActions, action => action.Reason == "BelowMinimumSize" || action.Reason == "NotSmaller");
        Assert.Equal(source.Length, result.Bytes.Length);
        Assert.Equal(source, result.Bytes);
    }

    [Fact]
    public void Optimize_RemovesUnreferencedObjects() {
        byte[] source = BuildPdfWithUnreferencedStream();

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source);

        Assert.True(result.Applied);
        Assert.Contains(result.Actions, action => action.Kind == "RemoveUnreferencedObject" && action.ObjectNumber == 6);
        Assert.True(result.OptimizedLengthBytes < result.OriginalLengthBytes);
        Assert.DoesNotContain("ORPHAN", PdfEncoding.Latin1GetString(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void ReachabilityWalk_HandlesDeepIndirectReferenceChainsIteratively() {
        const int objectCount = 25_000;
        var objects = new Dictionary<int, PdfIndirectObject>(objectCount);
        for (int objectNumber = 1; objectNumber <= objectCount; objectNumber++) {
            var dictionary = new PdfDictionary();
            if (objectNumber < objectCount) dictionary.Items["Next"] = new PdfReference(objectNumber + 1, 0);
            objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, dictionary);
        }

        var reachable = new HashSet<int>();
        MethodInfo method = typeof(PdfOptimizer).GetMethod(
            "CollectReachableObjectNumbers",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        method.Invoke(null, new object[] { objects, new PdfReference(1, 0), reachable });

        Assert.Equal(objectCount, reachable.Count);
    }

    [Fact]
    public void Optimize_DeduplicatesIdenticalStreamsAndRewritesReferences() {
        byte[] source = BuildPdfWithDuplicateStreams();

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, new PdfOptimizationOptions {
            CompressUnfilteredStreams = false,
            KeepOriginalWhenNotSmaller = false
        });

        Assert.True(result.Applied);
        Assert.Contains(result.Actions, action => action.Kind == "DeduplicateStream" && action.ObjectNumber == 6);
        string rewritten = PdfEncoding.Latin1GetString(result.Bytes);
        Assert.DoesNotContain("/Contents [5 0 R 6 0 R]", rewritten, StringComparison.Ordinal);
        Assert.Contains("5 0 R", rewritten, StringComparison.Ordinal);
        Assert.Contains("Duplicate", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Optimize_DeduplicatesStreamsUsingKeeperGeneration() {
        byte[] source = BuildPdfWithGeneratedDuplicateStreams();

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, new PdfOptimizationOptions {
            CompressUnfilteredStreams = false,
            KeepOriginalWhenNotSmaller = false
        });

        string rewritten = PdfEncoding.Latin1GetString(result.Bytes);

        Assert.True(result.Applied);
        Assert.Contains(result.Actions, action => action.Kind == "DeduplicateStream" && action.ObjectNumber == 6);
        Assert.Contains("5 0 R", rewritten, StringComparison.Ordinal);
        Assert.DoesNotContain("5 2 R", rewritten, StringComparison.Ordinal);
        Assert.Contains("Duplicate", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Optimize_RejectsSignedPdf() {
        byte[] signed = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 3 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "4 0 obj",
            "<< /FT /Sig /V 5 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Sig /ByteRange [0 0 0 0] /Contents <> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfOptimizer.Optimize(signed));
        Assert.Equal(PdfMutationOperation.Optimize, exception.Plan.Operation);
    }

    private static byte[] BuildPdfWithUncompressedTextStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] ExtractFirstStreamData(byte[] pdf) {
        byte[] streamMarker = Encoding.ASCII.GetBytes("stream\n");
        byte[] endMarker = Encoding.ASCII.GetBytes("\nendstream");
        int start = IndexOf(pdf, streamMarker, 0);
        Assert.True(start >= 0);
        start += streamMarker.Length;
        int end = IndexOf(pdf, endMarker, start);
        Assert.True(end > start);
        var data = new byte[end - start];
        Buffer.BlockCopy(pdf, start, data, 0, data.Length);
        return data;
    }

    private static int IndexOf(byte[] buffer, byte[] pattern, int startIndex) {
        for (int i = startIndex; i <= buffer.Length - pattern.Length; i++) {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++) {
                if (buffer[i + j] != pattern[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                return i;
            }
        }

        return -1;
    }

    private static byte[] BuildPdfWithUnreferencedStream() {
        string orphan = new string('O', 4096) + "ORPHAN";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length 34 >>",
            "stream",
            "BT /F1 12 Tf 72 720 Td (Text) Tj ET",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + orphan.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            orphan,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithDuplicateStreams() {
        string streamContent = "BT /F1 12 Tf 72 720 Td (" + new string('D', 1024) + " Duplicate) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + streamContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + streamContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithGeneratedDuplicateStreams() {
        string streamContent = "BT /F1 12 Tf 72 720 Td (" + new string('G', 1024) + " Duplicate) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents [5 0 R 6 2 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + streamContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "6 2 obj",
            "<< /Length " + streamContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
