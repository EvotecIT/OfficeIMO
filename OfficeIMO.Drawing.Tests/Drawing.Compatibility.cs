using OfficeIMO.Drawing;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingCompatibilityTests {
    private static readonly OfficeFormatDescriptor Doc = new(
        "Word.Doc",
        ".DOC",
        OfficeDocumentFamily.Word,
        OfficeDocumentKind.Document,
        OfficeFormatGeneration.Legacy,
        OfficeFormatEncoding.CompoundBinary,
        macroEnabled: true);

    private static readonly OfficeFormatDescriptor Docx = new(
        "Word.Docx",
        "docx",
        OfficeDocumentFamily.Word,
        OfficeDocumentKind.Document,
        OfficeFormatGeneration.Modern,
        OfficeFormatEncoding.OpenXml,
        macroEnabled: false);

    [Fact]
    public void DescriptorNormalizesExtensionAndUsesStableIdentity() {
        var equivalent = new OfficeFormatDescriptor(
            "Word.Doc",
            "doc",
            OfficeDocumentFamily.Word,
            OfficeDocumentKind.Document,
            OfficeFormatGeneration.Legacy,
            OfficeFormatEncoding.CompoundBinary,
            macroEnabled: true);

        Assert.Equal(".doc", Doc.Extension);
        Assert.Equal(Doc, equivalent);
        Assert.Equal(Doc.GetHashCode(), equivalent.GetHashCode());
    }

    [Fact]
    public void ReportKeepsIndependentFidelityDimensionsAndLossState() {
        var report = new OfficeCompatibilityReport(
            Doc,
            Docx,
            OfficeCompatibilityMode.PreferVisual,
            new[] {
                new OfficeCompatibilityFinding(
                    "Word.Chart.Rasterized",
                    "Charts",
                    "Chart was rendered as a picture.",
                    OfficeCompatibilityState.Rasterized,
                    OfficeCompatibilitySeverity.Warning,
                    OfficeCompatibilityImpact.Editability | OfficeCompatibilityImpact.Behavioral,
                    representsLoss: true,
                    sourceLocation: "body/chart1")
            });

        Assert.True(report.HasLoss);
        Assert.False(report.HasBlockedFeatures);
        Assert.False(report.IsStrictlyCompatible);
        Assert.Single(report.GetFindings(OfficeCompatibilityImpact.Behavioral));
        Assert.Throws<InvalidOperationException>(report.RequireNoLoss);
    }

    [Fact]
    public void ReportRejectsCrossFamilyFormats() {
        var xlsx = new OfficeFormatDescriptor(
            "Excel.Xlsx",
            ".xlsx",
            OfficeDocumentFamily.Excel,
            OfficeDocumentKind.Document,
            OfficeFormatGeneration.Modern,
            OfficeFormatEncoding.OpenXml,
            macroEnabled: false);

        Assert.Throws<ArgumentException>(() => new OfficeCompatibilityReport(
            Doc,
            xlsx,
            OfficeCompatibilityMode.StrictNative));
    }

    [Fact]
    public void SourceCarrierRoundTripsAndRejectsTamperedPackagePayload() {
        byte[] package = CreateMinimalOpcPackage();
        byte[] source = Encoding.UTF8.GetBytes("original source payload");

        byte[] carried = OfficeCompatibilitySourceCarrier.AttachToPackage(
            package,
            "Excel.Xls",
            "source.xls",
            OfficeCompatibilityMode.PreservationOnly,
            source);

        Assert.True(OfficeCompatibilitySourceCarrier.TryReadPackage(
            carried,
            out OfficeCompatibilitySourcePayload? payload,
            out string? error), error);
        Assert.NotNull(payload);
        Assert.Equal("Excel.Xls", payload!.FormatId);
        Assert.Equal("source.xls", payload.FileName);
        Assert.Equal(source, payload.ToArray());

        byte[] tampered = ReplacePackageEntry(
            carried,
            OfficeCompatibilitySourceCarrier.PayloadPath,
            Encoding.UTF8.GetBytes("tampered payload"));
        Assert.False(OfficeCompatibilitySourceCarrier.TryReadPackage(
            tampered,
            out payload,
            out error));
        Assert.Null(payload);
        Assert.Contains("SHA-256 mismatch", error, StringComparison.Ordinal);
    }

    private static byte[] CreateMinimalOpcPackage() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/></Types>");
            WriteEntry(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"document.xml\"/></Relationships>");
            WriteEntry(archive, "document.xml", "<document/>");
        }
        return output.ToArray();
    }

    private static byte[] ReplacePackageEntry(byte[] package, string path, byte[] replacement) {
        using var input = new MemoryStream(package, writable: false);
        using var source = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        using var output = new MemoryStream();
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                ZipArchiveEntry copied = target.CreateEntry(entry.FullName);
                using Stream destination = copied.Open();
                if (string.Equals(entry.FullName, path, StringComparison.OrdinalIgnoreCase)) {
                    destination.Write(replacement, 0, replacement.Length);
                } else {
                    using Stream original = entry.Open();
                    original.CopyTo(destination);
                }
            }
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string path, string text) {
        ZipArchiveEntry entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(text);
    }
}
