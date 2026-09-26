using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
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
    public void LockedSimpleFieldsBecomeFixedOdtFieldsExceptPageCount() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(
            new SimpleField(new Run(new Text("4"))) { Instruction = " PAGE ", FieldLock = true },
            new SimpleField(new Run(new Text("Today"))) { Instruction = " DATE ", FieldLock = true },
            new SimpleField(new Run(new Text("09:30"))) { Instruction = " TIME ", FieldLock = true },
            new SimpleField(new Run(new Text("12"))) { Instruction = " NUMPAGES ", FieldLock = true });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtParagraph result = Assert.Single(conversion.Value.Paragraphs);
        Assert.Equal(new[] { OdtFieldKind.PageNumber, OdtFieldKind.Date, OdtFieldKind.Time },
            result.Fields.Select(field => field.Kind));
        Assert.All(result.Fields, field => Assert.True(field.IsFixed));
        Assert.Equal("4Today09:3012", result.Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void FieldInsideUnmodeledInlineWrapperIsInspectedLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph();
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("content.xml").Descendants(text + "p").Single().Add(
            new XElement(text + "meta", new XElement(text + "page-number", "7")));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-text-fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void WordNoteFieldResultRetainsItsVisibleCachedText() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Before ");
        Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<Footnote>().Single(item => item.Type == null);
        note.Descendants<Paragraph>().First().Append(
            new SimpleField(new Run(new Text("7"))) { Instruction = " PAGE " },
            new Run(new Text(" after")));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtNote converted = Assert.Single(conversion.Value.Paragraphs.Single().Notes);
        Assert.Equal("Before 7 after", converted.Paragraphs.Single().Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void SimpleFieldInsideComplexInstructionIsNotConvertedOrDisplayed() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" IF ")),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Accepted")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtParagraph result = Assert.Single(conversion.Value.Paragraphs);
        Assert.Equal("Accepted", result.Text);
        Assert.Empty(result.Fields);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void ComplexInstructionAcrossParagraphsHidesNestedSimpleField() {
        using WordDocument source = WordDocument.Create();
        WordParagraph instruction = source.AddParagraph();
        instruction._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" IF ")));
        WordParagraph resultParagraph = source.AddParagraph();
        resultParagraph._paragraph.Append(
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Accepted")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        Assert.Same(instruction._paragraph.Ancestors().Last(), resultParagraph._paragraph.Ancestors().Last());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtParagraph result = conversion.Value.Paragraphs.Last();
        Assert.Equal("Accepted", result.Text);
        Assert.Empty(result.Fields);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void SimpleFieldInsideComplexResultRetainsCachedTextWithLoss() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" IF 1 = 1 ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Page ")),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        OdtParagraph result = Assert.Single(conversion.Value.Paragraphs);
        Assert.Equal("Page 5", result.Text);
        Assert.Empty(result.Fields);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
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
    public void OdtFieldUsesTransformedCachedTextAndReportsDynamicTransformLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.TextTransform = OdfTextTransform.Lowercase;
        paragraph.AddField(OdtFieldKind.Date, "TODAY");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Equal("today", Assert.Single(word.InspectFields()).ResultText);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "inline-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdtFieldRetainsParagraphRunFormattingInWordResult() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.Bold = true;
        paragraph.Color = OdfColor.Parse("#336699");
        paragraph.FontSize = OdfLength.Points(14);
        paragraph.AddField(OdtFieldKind.PageNumber, "3");

        using WordDocument word = source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Run run = word.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<SimpleField>().Single().Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.Equal("336699", run.RunProperties?.Color?.Val?.Value);
        Assert.Equal("28", run.RunProperties?.FontSize?.Val?.Value);
    }

    [Fact]
    public void FieldLocalNamespaceDeclarationDoesNotMakeBasicFieldUnsupported() {
        OdtDocument source = OdtDocument.Create();
        OdtField field = source.AddParagraph().AddField(OdtFieldKind.Date, "Today");
        field.Element.Add(new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "t",
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0"));
        source.Package.MarkXmlDirty("content.xml");

        Assert.True(field.IsBasic);
        using WordDocument word = source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Assert.Equal("Today", Assert.Single(word.InspectFields()).ResultText);
    }

    [Fact]
    public void FlattenedUnsupportedFieldRetainsEffectiveParagraphFormatting() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph();
        paragraph.Bold = true;
        paragraph.Color = OdfColor.Parse("#336699");
        paragraph.AddField(OdtFieldKind.PageNumber, "7");
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("content.xml").Descendants(text + "page-number").Single()
            .SetAttributeValue(text + "page-adjust", "1");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Run run = word.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Descendants<Run>()
            .Single(item => item.InnerText == "7");
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.Equal("336699", run.RunProperties?.Color?.Val?.Value);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
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
    public void FormattedWordFieldResultReportsFormattingLossAlongsideCachedText() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph.AddText("Page ");
        paragraph._paragraph.Append(new SimpleField(new Run(
            new RunProperties(new Bold(), new Color { Val = "336699" }),
            new Text("7"))) { Instruction = " PAGE " });
        paragraph.AddText(" of ");
        paragraph._paragraph.Append(new SimpleField(new Run(new Text("12"))) {
            Instruction = " NUMPAGES "
        });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Page 7 of 12", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "field-result-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
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
    public void FieldInNestedTableCellListIsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddTable(1, 1);
        var tableNamespace = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        var cell = source.Package.GetXml("content.xml").Descendants(tableNamespace + "table-cell").Single();
        cell.Add(new System.Xml.Linq.XElement(text + "list",
            new System.Xml.Linq.XElement(text + "list-item",
                new System.Xml.Linq.XElement(text + "p",
                    new System.Xml.Linq.XElement(text + "page-number", "2")))));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-text-fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
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
        OdtHyperlink fieldLink = Assert.Single(conversion.Value.Paragraphs.Single().Hyperlinks, link => link.Text == "5");
        Assert.Equal("#section", fieldLink.Href);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CustomXmlNestedFieldRetainsCachedTextAndReportsLoss() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new CustomXmlRun(
            new Run(new Text("Before ")),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " },
            new Run(new Text(" after"))));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Before 5 after", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ContentControlInsideSimpleFieldEmitsCachedTextOnce() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new Run(new Text("Page ")),
            new SimpleField(new SdtRun(new SdtContentRun(new Run(new Text("5"))))) {
                Instruction = " PAGE "
            });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Page 5", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void SimpleFieldInsideContentControlKeepsOrderAndReportsLoss() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new SdtRun(new SdtContentRun(
            new Run(new Text("Page ")),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " },
            new Run(new Text(" today")))));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Page 5 today", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void SimpleFieldInsideCustomXmlContentControlEmitsCachedTextOnce() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new CustomXmlRun(
            new SdtRun(new SdtContentRun(
                new Run(new Text("Page ")),
                new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " }))));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("Page 5", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ContentControlBeforeHyperlinkFieldPreservesInlineOrder() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new SdtRun(new SdtContentRun(new Run(new Text("A")))),
            new SimpleField(new Run(new Text("5"))) { Instruction = " PAGE " }) { Anchor = "section" });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Equal("A5", Assert.Single(conversion.Value.Paragraphs).Text);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ComplexFieldInstructionRunsAndNestedSimpleFieldStayHidden() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new Text("instruction")),
            new SimpleField(new Run(new Text("hidden"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Visible")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();

        Assert.Equal("Visible", Assert.Single(conversion.Value.Paragraphs).Text);
    }

    [Fact]
    public void ComplexFieldInstructionAcrossParagraphsStaysHiddenInOdt() {
        using WordDocument source = WordDocument.Create();
        WordParagraph start = source.AddParagraph("Start ");
        start._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        source.AddParagraph("Hidden instruction");
        WordParagraph result = source.AddParagraph("Result ");
        result._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Visible")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Value.Paragraphs, paragraph => paragraph.Text.Contains("Hidden instruction"));
        Assert.Contains(conversion.Value.Paragraphs, paragraph => paragraph.Text.Contains("Visible"));
    }

    [Fact]
    public void HandledHyperlinkFieldDoesNotHideInspectedDrawingField() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph();
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        XNamespace xlink = "http://www.w3.org/1999/xlink";
        XDocument content = source.Package.GetXml("content.xml");
        content.Descendants(text + "p").Single().Add(new XElement(text + "a",
            new XAttribute(xlink + "href", "https://example.com"),
            new XElement(text + "date", "Linked date")));
        content.Descendants(office + "text").Single().Add(new XElement(draw + "frame",
            new XElement(draw + "text-box", new XElement(text + "p",
                new XElement(text + "date", "Drawing date")))));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Hyperlink linkedField = Assert.Single(target._wordprocessingDocument.MainDocumentPart!.Document!.Body!
            .Descendants<Hyperlink>());
        Assert.Equal("https://example.com/", target._wordprocessingDocument.MainDocumentPart!
            .HyperlinkRelationships.Single(relationship => relationship.Id == linkedField.Id!.Value).Uri.ToString());
        Assert.Equal("Linked date", string.Concat(linkedField.Descendants<Text>().Select(text => text.Text)));
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-text-fields" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
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
        Assert.DoesNotContain(toWord.Report.Mappings, mapping => mapping.Feature == "source-text-fields");
        Assert.Empty(target.InspectFields());
        Assert.Throws<OdfConversionLossException>(() => odt.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
