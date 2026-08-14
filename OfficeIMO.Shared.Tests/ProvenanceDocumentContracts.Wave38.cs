using System.IO.Compression;
using System.Reflection;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightTreatsQuotesAsDataInEndTags() {
        AssertHtmlPreflightRejects("</div x=a\"><span></span><span></span>", maximumEntries: 1);
    }

    [Fact]
    public void HtmlPreflightConsumesPunctuationInTagNames() {
        AssertHtmlPreflightRejects("<script.foo><span></span><span></span>", maximumEntries: 1);
    }

    [Fact]
    public void ExcelXlsbRejectsSignatureOriginTargetsWithWorkbookContentType() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: true),
            "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "<Relationship Id=\"rSig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options));

        Assert.Contains("signature-origin target", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertHtmlPreflightRejects(string html, int maximumEntries) {
        MethodInfo method = typeof(HtmlProvenance).GetMethod(
            "ValidatePotentialElementCount",
            BindingFlags.NonPublic | BindingFlags.Static) ?? throw new MissingMethodException();
        TargetInvocationException exception = Assert.Throws<TargetInvocationException>(() =>
            method.Invoke(null, new object[] { html, maximumEntries }));
        Assert.IsType<InvalidDataException>(exception.InnerException);
    }

    private static byte[] ReplaceWave38Entry(byte[] package, string name, string content) {
        using var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        output.Position = 0;
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            archive.GetEntry(name)!.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(name, CompressionLevel.Optimal);
            using Stream stream = replacement.Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }
        return output.ToArray();
    }
}
