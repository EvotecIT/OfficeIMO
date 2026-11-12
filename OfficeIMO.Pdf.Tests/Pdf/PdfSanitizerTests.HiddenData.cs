using System.Text;
using OfficeIMO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfSanitizerTests {
    [Fact]
    public void HiddenDataReportsPdfCategoriesAndDefaultRemovalKeepsOptInKinds() {
        using PdfDocument document = PdfDocument.Load(BuildBeforeSharingPdf());

        OfficeHiddenDataReport inspection = document.InspectHiddenData();

        Assert.Equal("PDF", inspection.Format);
        Assert.All(inspection.Findings, static finding => Assert.Equal(OfficeHiddenDataOutcome.Found, finding.Outcome));
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.DocumentProperties),
            static finding => finding.Location == "Info/Author" && finding.Value == "PRIVATE-AUTHOR");
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.DocumentProperties), static finding => finding.Location == "Catalog/Metadata");
        Assert.Equal("PRIVATE-COMMENT", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Comments)).Value);
        Assert.Contains(inspection.GetFindings(OfficeHiddenDataKind.ActiveContent),
            static finding => finding.Description.Contains("JavaScript", StringComparison.Ordinal));
        Assert.Equal("payload.txt", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Attachments)).Value);
        Assert.Equal("PRIVATE-BOOKMARK", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.Bookmarks)).Value);
        Assert.Equal("PRIVATE-LAYER", Assert.Single(inspection.GetFindings(OfficeHiddenDataKind.HiddenContent)).Value);

        OfficeHiddenDataResult removal = document.RemoveHiddenData();

        foreach (OfficeHiddenDataKind removed in new[] { OfficeHiddenDataKind.DocumentProperties, OfficeHiddenDataKind.Comments, OfficeHiddenDataKind.ActiveContent }) {
            Assert.NotEmpty(removal.Report.GetFindings(removed));
            Assert.All(removal.Report.GetFindings(removed).Where(static finding => !IsKeptTitle(finding)),
                static finding => Assert.Equal(OfficeHiddenDataOutcome.Removed, finding.Outcome));
        }
        foreach (OfficeHiddenDataKind kept in new[] { OfficeHiddenDataKind.Attachments, OfficeHiddenDataKind.Bookmarks, OfficeHiddenDataKind.HiddenContent }) {
            OfficeHiddenDataFinding finding = Assert.Single(removal.Report.GetFindings(kept));
            Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome);
            Assert.NotNull(finding.Note);
        }

        using PdfDocument cleaned = PdfDocument.Load(removal.ToArray());
        OfficeHiddenDataReport after = cleaned.InspectHiddenData();
        Assert.All(after.GetFindings(OfficeHiddenDataKind.DocumentProperties), static finding => Assert.True(IsKeptTitle(finding)));
        Assert.Empty(after.GetFindings(OfficeHiddenDataKind.Comments));
        Assert.Empty(after.GetFindings(OfficeHiddenDataKind.ActiveContent));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.Attachments));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.Bookmarks));
        Assert.Single(after.GetFindings(OfficeHiddenDataKind.HiddenContent));
        Assert.Contains("VISIBLE-PAGE-CONTENT", cleaned.Read().Text, StringComparison.Ordinal);
        Assert.Equal(inspection.Findings.Count, document.InspectHiddenData().Findings.Count);
    }

    [Fact]
    public void DefaultHiddenDataRemovalKeepsTheTitleAndPdfAAndPdfUaIdentification() {
        using PdfDocument document = PdfDocument.Load(BuildConformingMetadataPdf());

        OfficeHiddenDataResult removal = document.RemoveHiddenData();

        string[] kept = { "Info/Title", "Catalog/Metadata/dc:title", "Catalog/Metadata/pdfaid", "Catalog/Metadata/pdfuaid" };
        foreach (string location in kept) {
            OfficeHiddenDataFinding finding = Assert.Single(removal.Report.Findings, item => item.Location == location);
            Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome);
            Assert.Equal(OfficeHiddenDataReason.NeededByDocument, finding.Reason);
            Assert.StartsWith("Kept because", finding.Note, StringComparison.Ordinal);
        }
        Assert.Equal("PDF/A-2B", Assert.Single(removal.Report.Findings, static item => item.Location == "Catalog/Metadata/pdfaid").Value);
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removal.Report.Findings, static item => item.Location == "Catalog/Metadata").Outcome);
        Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removal.Report.Findings, static item => item.Location == "Info/Author").Outcome);

        byte[] cleaned = removal.ToArray();
        PdfDocumentInfo info = PdfDocument.Load(cleaned).Inspect();
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.Equal("Quarterly report", info.Metadata.Title);
        Assert.Equal("Quarterly report", xmp.Title);
        Assert.Equal(2, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.Equal(1, xmp.PdfUaPart);
        Assert.Null(xmp.Filter);
        Assert.Contains("http://www.aiim.org/pdfua/ns/id/", xmp.RawXml, StringComparison.Ordinal);
        Assert.Equal(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero), xmp.CreationDate);
        Assert.DoesNotContain("PRIVATE-", PdfEncoding.Latin1GetString(cleaned), StringComparison.Ordinal);
        Assert.DoesNotContain(PdfDocument.Load(cleaned).RemoveHiddenData().Report.GetFindings(OfficeHiddenDataKind.DocumentProperties),
            static finding => finding.Outcome == OfficeHiddenDataOutcome.Removed);
    }

    private static bool IsKeptTitle(OfficeHiddenDataFinding finding) =>
        finding.Location is "Info/Title" or "Catalog/Metadata/dc:title" && finding.Outcome != OfficeHiddenDataOutcome.Removed;

    private static byte[] BuildConformingMetadataPdf() {
        const string xmp =
            "<?xpacket begin='' id='W5M0MpCehiHzreSzNTczkc9d'?><x:xmpmeta xmlns:x='adobe:ns:meta/' x:xmptk='PRIVATE-TOOLKIT'>" +
            "<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>" +
            "<rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:xmp='http://ns.adobe.com/xap/1.0/' " +
            "xmlns:pdf='http://ns.adobe.com/pdf/1.3/' xmlns:pdfaid='http://www.aiim.org/pdfa/ns/id/' xmlns:pdfuaid='http://www.aiim.org/pdfua/ns/id/' " +
            "xmlns:custom='urn:private:custom' pdfuaid:part='1' xmp:CreatorTool='PRIVATE-CREATOR-TOOL'>" +
            "<dc:title><rdf:Alt><rdf:li xml:lang='x-default'>Quarterly report</rdf:li></rdf:Alt></dc:title>" +
            "<dc:creator><rdf:Seq><rdf:li>PRIVATE-XMP-AUTHOR</rdf:li></rdf:Seq></dc:creator>" +
            "<xmp:CreateDate>2026-01-02T03:04:05Z</xmp:CreateDate><pdf:Producer>OfficeIMO-Test</pdf:Producer>" +
            "<custom:Owner>PRIVATE-CUSTOM-OWNER</custom:Owner>" +
            "<pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance></rdf:Description>" +
            "<rdf:Description rdf:about='' xmlns:pdfaExtension='http://www.aiim.org/pdfa/ns/extension/' " +
            "xmlns:pdfaSchema='http://www.aiim.org/pdfa/ns/schema#'><pdfaExtension:schemas><rdf:Bag>" +
            "<rdf:li rdf:parseType='Resource'><pdfaSchema:schema>PDF/UA identification</pdfaSchema:schema>" +
            "<pdfaSchema:namespaceURI>http://www.aiim.org/pdfua/ns/id/</pdfaSchema:namespaceURI><pdfaSchema:prefix>pdfuaid</pdfaSchema:prefix></rdf:li>" +
            "<rdf:li rdf:parseType='Resource'><pdfaSchema:schema>PRIVATE-SCHEMA</pdfaSchema:schema>" +
            "<pdfaSchema:namespaceURI>urn:private:custom</pdfaSchema:namespaceURI><pdfaSchema:prefix>custom</pdfaSchema:prefix></rdf:li>" +
            "</rdf:Bag></pdfaExtension:schemas></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end='w'?>";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R /ViewerPreferences << /DisplayDocTitle true >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            StreamObject(4, string.Empty, string.Empty),
            StreamObject(5, "/Type /Metadata /Subtype /XML", xmp),
            "6 0 obj", "<< /Title (Quarterly report) /Author (PRIVATE-AUTHOR) /Creator (PRIVATE-CREATOR) /Producer (OfficeIMO-Test) /CreationDate (D:20260102030405Z) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 6 0 R /Size 7 >>", "%%EOF"
        }) + "\n");
    }
}
