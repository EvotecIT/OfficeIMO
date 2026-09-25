using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtFieldConversionTests {
    [Fact]
    public void WordSimpleFieldsRetainInlineOrderAndCachedTextInOdt() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph.AddText("Page ");
        paragraph.AddField(WordFieldType.Page);
        paragraph.Field!.Text = "3";
        paragraph.AddText(" of ");
        paragraph.AddField(WordFieldType.NumPages);
        paragraph.Field!.Text = "12";

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        OdtParagraph result = Assert.Single(reopened.Paragraphs);
        Assert.Equal("Page 3 of 12", result.Text);
        Assert.Equal(new[] { OdtFieldKind.PageNumber, OdtFieldKind.PageCount },
            result.Fields.Select(field => field.Kind));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        source.ToOpenDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss });
    }

    [Fact]
    public void OdtFieldsBecomeWordSimpleFieldsWithCachedResults() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Dated ");
        paragraph.AddField(OdtFieldKind.Date, "2026-09-25");
        paragraph.AddText(" at ");
        paragraph.AddField(OdtFieldKind.Time, "09:30");
        OdtField fixedPage = paragraph.AddField(OdtFieldKind.PageNumber, "4");
        fixedPage.IsFixed = true;

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        using var stream = new MemoryStream();
        target.Save(stream);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordFieldInfo[] fields = reopened.InspectFields().ToArray();
        Assert.Equal(new WordFieldType?[] { WordFieldType.Date, WordFieldType.Time, WordFieldType.Page },
            fields.Select(field => field.FieldType));
        Assert.Equal(new[] { "2026-09-25", "09:30", "4" }, fields.Select(field => field.ResultText));
        Assert.True(fields[2].IsLocked);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 3);
        using WordDocument strict = source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
    }

    [Fact]
    public void OdtFieldWithSpacesStaysNativeAcrossConversion() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddField(OdtFieldKind.Date, "September 25, 2026");
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        Assert.Empty(source.Package.GetXml("content.xml").Descendants(text + "date").Single().Elements());

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss });
        using WordDocument word = conversion.Value;
        Assert.Equal("September 25, 2026", Assert.Single(word.InspectFields()).ResultText);
        Assert.Equal("September 25, 2026", Assert.Single(Assert.Single(source.Paragraphs).Fields).DisplayText);
    }

    [Fact]
    public void OdtXmlBooleanFixedValuesMapToWordLockState() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddField(OdtFieldKind.Date, "Today");
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        var date = source.Package.GetXml("content.xml").Descendants(text + "date").Single();
        date.SetAttributeValue(text + "fixed", "1");
        source.Package.MarkXmlDirty("content.xml");

        using WordDocument word = source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Assert.True(Assert.Single(word.InspectFields()).IsLocked);

        date.SetAttributeValue(text + "fixed", "0");
        using WordDocument dynamicWord = source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Assert.False(Assert.Single(dynamicWord.InspectFields()).IsLocked);
    }

    [Fact]
    public void OdtCachedFieldBoundarySpacesSurviveWordSaveAndReopen() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddField(OdtFieldKind.Date, "  September 25  ");

        using WordDocument word = source.ToWordDocumentResult().Value;
        using var output = new MemoryStream();
        word.Save(output);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(output.ToArray()));
        Assert.Equal("  September 25  ", Assert.Single(reopened.InspectFields()).ResultText);
    }

    [Fact]
    public void WordFieldResultTabAndBreakRetainVisibleTextWithLoss() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new SimpleField(new Run(
            new Text("A"), new TabChar(), new Text("B"), new Break(), new Text("C"))) {
            Instruction = " PAGE "
        });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("A\tB\nC", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void RepeatedMappedTableFieldsDoNotCancelUnsupportedFieldLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtTable table = source.AddTable(1, 1);
        table.Cell(0, 0).Paragraphs[0].AddField(OdtFieldKind.PageNumber, "1");
        source.AddParagraph().AddField(OdtFieldKind.Date, "2026-09-25");
        var tableNamespace = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        var content = source.Package.GetXml("content.xml");
        content.Descendants(tableNamespace + "table-row").Single()
            .SetAttributeValue(tableNamespace + "number-rows-repeated", "2");
        content.Descendants(text + "date").Single().SetAttributeValue(text + "date-value", "2026-09-25");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Equal(2, word.InspectFields().Count);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void FieldInUnconvertedNestedTableIsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddTable(1, 1);
        var tableNamespace = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        var cell = source.Package.GetXml("content.xml").Descendants(tableNamespace + "table-cell").Single();
        cell.Add(new System.Xml.Linq.XElement(tableNamespace + "table",
            new System.Xml.Linq.XElement(tableNamespace + "table-row",
                new System.Xml.Linq.XElement(tableNamespace + "table-cell",
                    new System.Xml.Linq.XElement(text + "p",
                        new System.Xml.Linq.XElement(text + "page-number", "2"))))));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-text-fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void HyperlinkNestedFieldRetainsCachedTextAndReportsLoss() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new Run(new Text("See ")),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " }) { Anchor = "section" });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("See 5", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnsupportedWordInstructionsAndOdtAdjustmentsReportLoss() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddField(WordFieldType.Ref);
        paragraph.Field!.Text = "Reference text";
        OdfConversionResult<OdtDocument> toOdt = word.ToOpenDocumentResult();
        Assert.Equal("Reference text", Assert.Single(toOdt.Value.Paragraphs).Text);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => word.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));

        OdtDocument odt = OdtDocument.Create();
        odt.AddParagraph().AddField(OdtFieldKind.PageNumber, "7");
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        odt.Package.GetXml("content.xml").Descendants(text + "page-number").Single()
            .SetAttributeValue(text + "page-adjust", "1");
        odt.Package.MarkXmlDirty("content.xml");
        OdfConversionResult<WordDocument> toWord = odt.ToWordDocumentResult();
        using WordDocument target = toWord.Value;
        Assert.Contains(toWord.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Empty(target.InspectFields());
        Assert.Throws<OdfConversionLossException>(() => odt.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
