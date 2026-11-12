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

    [Fact]
    public void DefaultHiddenDataRemovalKeepsPdfXAndInvoiceIdentificationWhileTheInvoiceStaysAttached() {
        byte[] source = BuildPdfXInvoiceMetadataPdf();

        OfficeHiddenDataResult removal = PdfDocument.Load(source).RemoveHiddenData();

        foreach (string location in new[] { "Info/GTS_PDFXVersion", "Catalog/Metadata/pdfxid", "Catalog/Metadata/xmpMM", "Catalog/Metadata/invoice" }) {
            OfficeHiddenDataFinding finding = Assert.Single(removal.Report.Findings, item => item.Location == location);
            Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome);
            Assert.Equal(OfficeHiddenDataReason.NeededByDocument, finding.Reason);
        }
        PdfDocumentInfo info = PdfDocument.Load(removal.ToArray()).Inspect();
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.Equal("PDF/X-4", info.Metadata.PdfXVersion);
        Assert.Equal("PDF/X-4", xmp.PdfXVersion);
        Assert.Equal("uuid:00000000-0000-0000-0000-000000000001", xmp.DocumentId);
        Assert.Equal("default", xmp.RenditionClass);
        Assert.NotNull(xmp.MetadataDate);
        Assert.Equal("EN 16931", xmp.ElectronicInvoiceConformanceLevel);
        Assert.Equal("factur-x.xml", xmp.ElectronicInvoiceDocumentFileName);
        Assert.Contains(PdfElectronicInvoiceMetadata.FacturXNamespaceUri, xmp.RawXml, StringComparison.Ordinal);
        Assert.Single(info.Attachments);
        Assert.DoesNotContain("PRIVATE-", PdfEncoding.Latin1GetString(removal.ToArray()), StringComparison.Ordinal);

        OfficeHiddenDataResult withAttachments = PdfDocument.Load(source).RemoveHiddenData(new OfficeHiddenDataOptions {
            Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Attachments
        });
        PdfXmpMetadataInfo stripped = Assert.IsType<PdfXmpMetadataInfo>(PdfDocument.Load(withAttachments.ToArray()).Inspect().XmpMetadata);
        Assert.Null(stripped.ElectronicInvoiceDocumentFileName);
        Assert.DoesNotContain(withAttachments.Report.Findings, static finding => finding.Location == "Catalog/Metadata/invoice");
        Assert.Equal("PDF/X-4", stripped.PdfXVersion);
    }

    [Fact]
    public void DefaultHiddenDataRemovalCleansATaggedPdfAndPrunesOnlyTheRemovedCommentFromItsStructure() {
        byte[] source = BuildTaggedCommentPdf();

        OfficeHiddenDataResult removal = PdfDocument.Load(source).RemoveHiddenData();

        Assert.Equal(OfficeHiddenDataOutcome.Removed,
            Assert.Single(removal.Report.GetFindings(OfficeHiddenDataKind.Comments)).Outcome);
        PdfDocumentInfo info = PdfDocument.Load(removal.ToArray()).Inspect();
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.True(tagged.Marked);
        Assert.Equal("en-US", info.CatalogLanguage);
        Assert.True(info.ViewerPreferences?.GetBoolean("DisplayDocTitle"));
        Assert.Equal(new[] { "Document", "P" }, tagged.StructureTypes);
        Assert.Equal(1, tagged.MarkedContentReferenceCount);
        Assert.Equal(0, tagged.ObjectReferenceCount);
        Assert.Equal(new[] { 0 }, tagged.ParentTreeStructParentIndexes);
        Assert.Equal("Visible", tagged.RoleMap["Para"]);
        Assert.Empty(info.Annotations);
        Assert.DoesNotContain("PRIVATE-", PdfEncoding.Latin1GetString(removal.ToArray()), StringComparison.Ordinal);
    }

    private static bool IsKeptTitle(OfficeHiddenDataFinding finding) =>
        finding.Location is "Info/Title" or "Catalog/Metadata/dc:title" && finding.Outcome != OfficeHiddenDataOutcome.Removed;

    private static byte[] BuildPdfXInvoiceMetadataPdf() {
        const string xmp =
            "<?xpacket begin='' id='W5M0MpCehiHzreSzNTczkc9d'?><x:xmpmeta xmlns:x='adobe:ns:meta/'>" +
            "<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>" +
            "<rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:xmp='http://ns.adobe.com/xap/1.0/' " +
            "xmlns:xmpMM='http://ns.adobe.com/xap/1.0/mm/' xmlns:pdfxid='http://www.npes.org/pdfx/ns/id/' " +
            "xmlns:fx='urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#' xmlns:pdfaid='http://www.aiim.org/pdfa/ns/id/'>" +
            "<dc:title><rdf:Alt><rdf:li xml:lang='x-default'>Invoice 42</rdf:li></rdf:Alt></dc:title>" +
            "<dc:creator><rdf:Seq><rdf:li>PRIVATE-XMP-AUTHOR</rdf:li></rdf:Seq></dc:creator><xmp:CreatorTool>PRIVATE-CREATOR-TOOL</xmp:CreatorTool>" +
            "<xmp:CreateDate>2026-01-02T03:04:05Z</xmp:CreateDate><xmp:ModifyDate>2026-01-02T03:04:05Z</xmp:ModifyDate>" +
            "<xmp:MetadataDate>2026-01-02T03:04:05Z</xmp:MetadataDate>" +
            "<xmpMM:DocumentID>uuid:00000000-0000-0000-0000-000000000001</xmpMM:DocumentID>" +
            "<xmpMM:InstanceID>uuid:00000000-0000-0000-0000-000000000002</xmpMM:InstanceID>" +
            "<xmpMM:VersionID>1</xmpMM:VersionID><xmpMM:RenditionClass>default</xmpMM:RenditionClass>" +
            "<pdfxid:GTS_PDFXVersion>PDF/X-4</pdfxid:GTS_PDFXVersion><pdfaid:part>3</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>" +
            "<fx:DocumentType>INVOICE</fx:DocumentType><fx:DocumentFileName>factur-x.xml</fx:DocumentFileName>" +
            "<fx:Version>1.0</fx:Version><fx:ConformanceLevel>EN 16931</fx:ConformanceLevel></rdf:Description>" +
            "<rdf:Description rdf:about='' xmlns:pdfaExtension='http://www.aiim.org/pdfa/ns/extension/' " +
            "xmlns:pdfaSchema='http://www.aiim.org/pdfa/ns/schema#'><pdfaExtension:schemas><rdf:Bag>" +
            "<rdf:li rdf:parseType='Resource'><pdfaSchema:schema>Factur-X PDF/A Extension Schema</pdfaSchema:schema>" +
            "<pdfaSchema:namespaceURI>urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#</pdfaSchema:namespaceURI><pdfaSchema:prefix>fx</pdfaSchema:prefix></rdf:li>" +
            "</rdf:Bag></pdfaExtension:schemas></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end='w'?>";
        const string invoice = "<rsm:CrossIndustryInvoice xmlns:rsm='urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100'/>";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R /Names << /EmbeddedFiles << /Names [(factur-x.xml) 8 0 R] >> >> /AF [8 0 R] >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            StreamObject(4, string.Empty, string.Empty),
            StreamObject(5, "/Type /Metadata /Subtype /XML", xmp),
            "6 0 obj", "<< /Title (Invoice 42) /Author (PRIVATE-AUTHOR) /GTS_PDFXVersion (PDF/X-4) /Trapped /False /CreationDate (D:20260102030405Z) /ModDate (D:20260102030405Z) >>", "endobj",
            StreamObject(7, "/Type /EmbeddedFile /Subtype /text#2Fxml", invoice),
            "8 0 obj", "<< /Type /Filespec /F (factur-x.xml) /UF (factur-x.xml) /AFRelationship /Data /EF << /F 7 0 R /UF 7 0 R >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 6 0 R /Size 9 >>", "%%EOF"
        }) + "\n");
    }

    private static byte[] BuildTaggedCommentPdf() {
        const string content = "/P <</MCID 0>> BDC BT /F1 12 Tf 20 150 Td (VISIBLE-TAGGED-TEXT) Tj ET EMC";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 9 0 R /MarkInfo << /Marked true >> /Lang (en-US) /ViewerPreferences << /DisplayDocTitle true >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [6 0 R] /StructParents 0 /Tabs /S >>", "endobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (PRIVATE-TAGGED-COMMENT) /StructParent 1 >>", "endobj",
            "7 0 obj", "<< /Type /StructElem /S /P /P 8 0 R /Pg 3 0 R /K 0 >>", "endobj",
            "8 0 obj", "<< /Type /StructElem /S /Document /P 9 0 R /K [7 0 R 10 0 R] >>", "endobj",
            "9 0 obj", "<< /Type /StructTreeRoot /K 8 0 R /ParentTree 11 0 R /ParentTreeNextKey 2 /RoleMap << /Para /Visible >> >>", "endobj",
            "10 0 obj", "<< /Type /StructElem /S /Annot /P 8 0 R /Pg 3 0 R /K << /Type /OBJR /Obj 6 0 R /Pg 3 0 R >> >>", "endobj",
            "11 0 obj", "<< /Nums [0 [7 0 R] 1 10 0 R] >>", "endobj",
            "12 0 obj", "<< /Title (Tagged) /Author (PRIVATE-AUTHOR) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 12 0 R /Size 13 >>", "%%EOF"
        }) + "\n");
    }

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
