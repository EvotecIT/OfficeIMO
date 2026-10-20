using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserWordPdfMergeTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WordConversionCanBeMergedWithAnAdditionalPdfWithoutLosingHeadingsOrPageOwnership(bool convertedFirst) {
        WorkflowArtifacts workflow = CreateWorkflow(convertedFirst);
        PdfReadDocument converted = PdfReadDocument.Open(workflow.Converted);
        Assert.Equal(2, converted.Pages.Count);
        Assert.True(converted.HasTaggedContent);
        PdfTaggedContentInfo sourceTags = Assert.IsType<PdfTaggedContentInfo>(converted.TaggedContent);
        Assert.Contains("H1", sourceTags.StructureTypes);
        Assert.Contains("H2", sourceTags.StructureTypes);

        PdfReadDocument merged = PdfReadDocument.Open(workflow.Merged.Artifact.Bytes);
        Assert.Equal(3, merged.Pages.Count);
        int offset = convertedFirst ? 0 : 1;
        Assert.Contains("Delivery status", merged.Pages[offset].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("The review is ready for acceptance.", merged.Pages[offset].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Acceptance checklist", merged.Pages[offset + 1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("The owner confirms the delivery.", merged.Pages[offset + 1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Additional approval record", merged.Pages[convertedFirst ? 2 : 0].ExtractText(), StringComparison.Ordinal);

        PdfTaggedContentInfo tags = Assert.IsType<PdfTaggedContentInfo>(merged.TaggedContent);
        Assert.Single(tags.StructureElements, element => element.StructureType == "Document");
        PdfStructureElementInfo firstHeading = Assert.Single(tags.StructureElements, element => element.StructureType == "H1");
        PdfStructureElementInfo secondHeading = Assert.Single(tags.StructureElements, element => element.StructureType == "H2");
        Assert.All(firstHeading.MarkedContentReferences, reference => Assert.Equal(merged.Pages[offset].ObjectNumber, reference.PageObjectNumber));
        Assert.All(secondHeading.MarkedContentReferences, reference => Assert.Equal(merged.Pages[offset + 1].ObjectNumber, reference.PageObjectNumber));
        Assert.NotEmpty(firstHeading.MarkedContentReferences);
        Assert.NotEmpty(secondHeading.MarkedContentReferences);
        int appendixObject = merged.Pages[convertedFirst ? 2 : 0].ObjectNumber;
        Assert.DoesNotContain(tags.StructureElements.SelectMany(element => element.MarkedContentReferences), reference => reference.PageObjectNumber == appendixObject);

        PdfDocumentReadResult logical = PdfDocument.Load(workflow.Merged.Artifact.Bytes).Read();
        Assert.Contains(logical.Headings, heading => heading.Text == "Delivery status" && heading.Level == 1 && heading.PageNumber == offset + 1);
        Assert.Contains(logical.Headings, heading => heading.Text == "Acceptance checklist" && heading.Level == 2 && heading.PageNumber == offset + 2);
        Assert.Equal(3, workflow.Merged.PageCount);
        Assert.Equal("application/pdf", workflow.Merged.Artifact.ContentType);
        using JsonDocument report = JsonDocument.Parse(workflow.Merged.Report!.Bytes);
        Assert.Equal("3", report.RootElement.GetProperty("details").GetProperty("outputPages").GetString());
        Assert.Equal("2", report.RootElement.GetProperty("details").GetProperty("sourceCount").GetString());
        Assert.Equal(workflow.ConvertedSnapshot, workflow.Converted);
        Assert.Equal(workflow.AdditionalSnapshot, workflow.Additional);
    }

    internal static WorkflowArtifacts CreateWorkflow(bool convertedFirst) {
        using WordDocument word = WordDocument.Create();
        word.Settings.Language = "en-US";
        word.AddParagraph("Delivery status").Style = WordParagraphStyles.Heading1;
        word.AddParagraph("The review is ready for acceptance.");
        word.AddPageBreak();
        word.AddParagraph("Acceptance checklist").Style = WordParagraphStyles.Heading2;
        word.AddParagraph("The owner confirms the delivery.");
        byte[] docx = word.ToBytes();
        ConversionResult conversion = new BrowserConversionService().ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            new SelectedDocument("delivery.docx", ".docx", "DOCX", docx.LongLength, docx),
            limitExcelRows: false,
            BrowserPdfProfileCatalog.Faithful);
        byte[] converted = conversion.Bytes;
        byte[] additional = PdfDocument.Create(pdf => pdf.Content(content => content
            .Paragraph(paragraph => paragraph.Text("Additional approval record")))).ToBytes();
        byte[] convertedSnapshot = (byte[])converted.Clone();
        byte[] additionalSnapshot = (byte[])additional.Clone();
        SelectedDocument[] files = convertedFirst
            ? [PdfFile(converted, "delivery.pdf"), PdfFile(additional, "approval.pdf")]
            : [PdfFile(additional, "approval.pdf"), PdfFile(converted, "delivery.pdf")];
        PdfToolResult merged = new BrowserPdfToolService().Execute(new PdfToolRequest(
            PdfToolCatalog.Find("merge"), files, "all", 1, 90, PdfOptimizationProfile.Balanced,
            string.Empty, string.Empty, string.Empty, false));
        return new WorkflowArtifacts(docx, converted, additional, merged, convertedSnapshot, additionalSnapshot);
    }

    private static SelectedDocument PdfFile(byte[] bytes, string name) => new(name, ".pdf", "PDF", bytes.LongLength, bytes);

    internal sealed record WorkflowArtifacts(byte[] Docx, byte[] Converted, byte[] Additional, PdfToolResult Merged, byte[] ConvertedSnapshot, byte[] AdditionalSnapshot);
}
