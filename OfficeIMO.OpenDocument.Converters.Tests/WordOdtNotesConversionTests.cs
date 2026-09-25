using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtNotesConversionTests {
    [Fact]
    public void NumericCustomCitationLabelStillReportsLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Note");
        var text = (XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("content.xml").Descendants(text + "note-citation").Single()
            .SetAttributeValue(text + "label", "1");
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-citations" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void ParagraphStyleRunFormattingOnWordNoteAnchorReportsLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Styles styles = source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Bold())) {
            Type = W.StyleValues.Paragraph, StyleId = "StyledNoteAnchor"
        });
        W.Paragraph anchor = source.OpenXmlDocument.MainDocumentPart.Document!.Body!
            .Descendants<W.Paragraph>().First(paragraph => paragraph.Descendants<W.FootnoteReference>().Any());
        anchor.ParagraphProperties ??= new W.ParagraphProperties();
        anchor.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = "StyledNoteAnchor" };

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnreferencedWordNoteDefinitionsRemainExplicitLoss(bool endnote) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Anchor");
        if (endnote) {
            paragraph.AddEndNote("Unreferenced");
            source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
                .Descendants<W.EndnoteReference>().Single().Remove();
        } else {
            paragraph.AddFootNote("Referenced");
            source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Append(
                new W.Footnote(new W.Paragraph(new W.Run(new W.Text("Unreferenced")))) { Id = 99 });
        }

        OdfConversionResult<OdtDocument> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == (endnote ? "source-endnotes" : "source-footnotes") &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void RepeatedReferenceDoesNotHideAnUnreferencedFootnoteDefinition() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Referenced");
        W.Paragraph paragraph = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.Paragraph>().First(item => item.Descendants<W.FootnoteReference>().Any());
        W.FootnoteReference reference = paragraph.Descendants<W.FootnoteReference>().Single();
        paragraph.Append(new W.Run((W.FootnoteReference)reference.CloneNode(true)));
        source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Append(
            new W.Footnote(new W.Paragraph(new W.Run(new W.Text("Unreferenced")))) { Id = 99 });

        OdfConversionResult<OdtDocument> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void DuplicateWordFootnoteDefinitionRemainsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("First");
        W.Footnote original = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(note => note.Type == null);
        source.OpenXmlDocument.MainDocumentPart.FootnotesPart.Footnotes.Append(
            new W.Footnote(new W.Paragraph(new W.Run(new W.Text("Second")))) { Id = original.Id!.Value });

        OdfConversionResult<OdtDocument> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-footnotes" && mapping.Status == OdfConversionMappingStatus.Unsupported &&
            mapping.Count == 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordNoteReferenceWithoutRunPropertiesIsConverted(bool endnote) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Anchor");
        if (endnote) paragraph.AddEndNote("Note body");
        else paragraph.AddFootNote("Note body");
        W.Run anchor = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.Run>().Single(run => endnote
                ? run.GetFirstChild<W.EndnoteReference>() != null
                : run.GetFirstChild<W.FootnoteReference>() != null);
        anchor.RunProperties?.Remove();

        OdtDocument converted = source.ToOpenDocument();
        OdtNote note = Assert.Single(Assert.Single(converted.Paragraphs).Notes);
        Assert.Equal(endnote ? OdtNoteKind.Endnote : OdtNoteKind.Footnote, note.Kind);
        Assert.Equal("Note body", note.Paragraphs[0].Text);
    }

    [Fact]
    public void WordNoteInlineSymbolIsExplicitContentLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        note.Elements<W.Paragraph>().Single().Append(new W.Run(new W.SymbolChar {
            Font = "Wingdings", Char = "F0A7"
        }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteInlineFieldWrapperIsExplicitContentLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        note.Elements<W.Paragraph>().Single().Append(new W.SimpleField(
            new W.Run(new W.Text("Field result"))) { Instruction = "DATE" });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordDocumentDefaultNoteFormattingIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new W.Bold());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CharacterStyleOnWordNoteBodyIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        W.Run body = note.Descendants<W.Run>().First(run => run.Descendants<W.Text>().Any());
        body.RunProperties ??= new W.RunProperties();
        body.RunProperties.RunStyle = new W.RunStyle { Val = "Emphasis" };

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingOrRepeatedWordNoteReferenceMarkIsExplicitLoss(bool repeated) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        W.FootnoteReferenceMark mark = note.Descendants<W.FootnoteReferenceMark>().Single();
        if (repeated) mark.Parent!.Append((W.FootnoteReferenceMark)mark.CloneNode(true));
        else mark.Remove();

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteReferenceMarkAfterBodyTextIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        W.Run markRun = note.Descendants<W.Run>().Single(run => run.Descendants<W.FootnoteReferenceMark>().Any());
        markRun.Remove();
        note.Elements<W.Paragraph>().Single().Append(markRun);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void StyledTextSharingWordNoteReferenceRunIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        W.Run markRun = note.Descendants<W.Run>().Single(run => run.Descendants<W.FootnoteReferenceMark>().Any());
        markRun.Append(new W.Text("Styled text"));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void ChangedWordDocumentDefaultSizeIsExplicitNoteLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!
            .RunPropertiesDefault!.RunPropertiesBaseStyle!.GetFirstChild<W.FontSize>()!.Val = "30";

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteTabAndWrappingBreakRemainSupportedText() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Before");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        note.Elements<W.Paragraph>().Single().Append(new W.Run(new W.TabChar(),
            new W.Text("Middle"), new W.Break(), new W.Text("After")));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains("\tMiddle\nAfter", Assert.Single(Assert.Single(conversion.Value.Paragraphs).Notes).Paragraphs[0].Text);
    }

    [Fact]
    public void NestedOdtNoteIsVisibleFromItsContainingParagraph() {
        OdtDocument source = OdtDocument.Create();
        OdtNote outer = source.AddParagraph("Anchor").AddFootnote("Outer body");
        OdtParagraph innerParagraph = outer.Paragraphs[0];
        OdtNote nested = innerParagraph.AddFootnote("Inner body");
        Assert.Equal(nested.Id, Assert.Single(innerParagraph.Notes).Id);

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        OdtNote reopenedOuter = Assert.Single(Assert.Single(reopened.Paragraphs).Notes);
        Assert.Equal("Inner body", Assert.Single(reopenedOuter.Paragraphs[0].Notes).Paragraphs[0].Text);
    }

    [Fact]
    public void AdditionalFootnoteReferenceInOneRunHasExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("First body");
        source.AddParagraph("Other").AddFootNote("Second body");
        W.FootnoteReference second = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.FootnoteReference>().Last();
        W.Run firstRun = source.OpenXmlDocument.MainDocumentPart.Document.Body
            .Descendants<W.Run>().First(run => run.GetFirstChild<W.FootnoteReference>() != null);
        firstRun.Append((W.FootnoteReference)second.CloneNode(true));
        second.Remove();
        Assert.Equal(2, source.InspectFeatures().Features.Single(feature => feature.Name == "Footnotes").Count);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void EmptyWordNoteBodyParagraphIsPreserved() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("First paragraph");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        note.Append(new W.Paragraph());

        OdtDocument converted = source.ToOpenDocument();
        OdtNote odtNote = Assert.Single(Assert.Single(converted.Paragraphs).Notes);
        Assert.Equal(new[] { "First paragraph", "" }, odtNote.Paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public void DefaultOdtParagraphStyleOnNoteBodyHasExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Body");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "default-style", new XAttribute(style + "family", "paragraph"),
                new XElement(style + "paragraph-properties", new XAttribute(fo + "text-align", "center"))));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void DefaultOdtTextStyleOnNoteAnchorHasReferenceFormattingLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Body");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(style + "default-style", new XAttribute(style + "family", "paragraph"),
                new XElement(style + "text-properties", new XAttribute(fo + "font-weight", "bold"))));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteRunSemanticsAndPageBreaksAreExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Paragraph noteParagraph = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId)
            .Elements<W.Paragraph>().Single();
        W.Run run = noteParagraph.Descendants<W.Run>().First(item => item.GetFirstChild<W.Text>() != null);
        run.RunProperties ??= new W.RunProperties();
        run.RunProperties.Append(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript });
        run.RunProperties.Append(new W.SmallCaps());
        run.RunProperties.Append(new W.Shading { Val = W.ShadingPatternValues.DiagonalCross });
        run.Append(new W.Break { Type = W.BreakValues.Page });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void StyledWordNoteReferenceMarkIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        W.Run mark = note.Descendants<W.Run>().Single(run => run.GetFirstChild<W.FootnoteReferenceMark>() != null);
        mark.RunProperties!.Append(new W.Color { Val = "AA0000" });
        Assert.Empty(source.ValidateDocument());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void StyledWordNoteAnchorIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Run anchor = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.Run>().Single(run => run.GetFirstChild<W.FootnoteReference>() != null);
        anchor.RunProperties ??= new W.RunProperties();
        anchor.RunProperties.Append(new W.Color { Val = "AA0000" });
        Assert.Empty(source.ValidateDocument());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void InheritedOdtNoteReferenceFormattingIsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Anchor");
        paragraph.Bold = true;
        paragraph.AddFootnote("Body");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void NestedOdtSpanKeepsOuterNoteReferenceStyleLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Body");
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        XElement note = source.Package.GetXml("content.xml").Descendants(text + "note").Single();
        note.ReplaceWith(new XElement(text + "span", new XAttribute(text + "style-name", "OuterStyle"),
            new XElement(text + "span", note)));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordCustomMarkAndCustomizedSeparatorAreExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.FootnoteReference reference = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.FootnoteReference>().Single();
        reference.CustomMarkFollows = true;
        W.Footnote separator = source.OpenXmlDocument.MainDocumentPart.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(note => note.Type?.Value == W.FootnoteEndnoteValues.Separator);
        separator.Descendants<W.Run>().Single().Append(new W.Text("custom rule"));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-citations" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-separators" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void AlternateWordSeparatorReferenceIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        source.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Append(
            new W.Footnote(new W.Paragraph(new W.Run(new W.SeparatorMark()))) {
                Type = W.FootnoteEndnoteValues.Separator, Id = -2
            });
        W.Settings settings = source.OpenXmlDocument.MainDocumentPart.DocumentSettingsPart!.Settings!;
        settings.AddChild(new W.FootnoteDocumentWideProperties(
            new W.NumberingFormat { Val = W.NumberFormatValues.UpperRoman },
            new W.FootnoteSpecialReference { Id = -2 }), true);
        Assert.Empty(source.ValidateDocument());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-separators" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-numbering-placement" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void StyledOdtReferenceAndCitationLabelAreExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Anchor");
        paragraph.AddFootnote("Body");
        byte[] modified = OdfTestPackageRewriter.Rewrite(source.ToBytes(), (name, bytes) => {
            if (name != "content.xml") return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XElement note = xml.Descendants(text + "note").Single();
            note.Element(text + "note-citation")!.SetAttributeValue(text + "label", "*");
            note.ReplaceWith(new XElement(text + "span", new XAttribute(text + "style-name", "ReferenceStyle"), note));
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(modified));
        Assert.Equal("*", Assert.Single(Assert.Single(reopened.Paragraphs).Notes).Citation);

        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-citations" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdtNoteInsertionWorksWithoutOptionalStylesPart() {
        OdtDocument source = OdtDocument.Create();
        source.Package.RemoveEntry("styles.xml");
        source.AddParagraph("Anchor").AddFootnote("Body");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal("Body", Assert.Single(Assert.Single(reopened.Paragraphs).Notes).Paragraphs[0].Text);
    }

    [Fact]
    public void OdtHeaderNotesAreOmittedFromWordWithExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Header anchor").AddFootnote("Header note");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-headers-footers" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Empty(word.ValidateDocument());
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdtNoteInventoryAcrossContentAndStylesPartsIsConsumedOnce() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body anchor").AddFootnote("Body note");
        source.PageLayout.Header.AddParagraph("Header anchor").AddFootnote("Nested header note");
        XElement header = source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "header").Single();
        XElement paragraph = header.Element(OdfNamespaces.Text + "p")!;
        paragraph.Remove();
        header.Add(new XElement(OdfNamespaces.Table + "table",
            new XElement(OdfNamespaces.Table + "table-row",
                new XElement(OdfNamespaces.Table + "table-cell", paragraph))));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "source-text-notes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdtFootnoteSeparatorConfigurationIsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Body anchor").AddFootnote("Body note");
        _ = source.PageLayout;
        XElement properties = source.Package.GetXml("styles.xml")
            .Descendants(OdfNamespaces.Style + "page-layout-properties").First();
        properties.Add(new XElement(OdfNamespaces.Style + "footnote-sep",
            new XAttribute(OdfNamespaces.Style + "width", "0.02in")));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-configuration" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteBodyTablesAreReportedAsUnsupported() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Note paragraph");
        source.AddTable(1, 1);
        W.Table table = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Last();
        table.Remove();
        W.Footnote note = source.OpenXmlDocument.MainDocumentPart.FootnotesPart!.Footnotes!
            .Elements<W.Footnote>().Single(item => item.Id?.Value == source.FootNotes[0].ReferenceId);
        note.Append(table);
        Assert.Empty(source.ValidateDocument());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void DirectWordNoteParagraphSpacingIsReported() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Note paragraph");
        var noteParagraphs = source.FootNotes[0].Paragraphs;
        Assert.NotNull(noteParagraphs);
        noteParagraphs![0].LineSpacingAfterPoints = 8;

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void OdtNoteConfigurationIsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Body");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new XElement(text + "notes-configuration", new XAttribute(text + "note-class", "footnote"),
                new XAttribute(text + "start-value", "5")));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-configuration" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteSettingsInAnotherSectionRemainExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("First section").AddFootNote("First note");
        source.AddSection().AddFootnoteProperties(numberingFormat: WordNumberFormat.UpperRoman);
        source.Sections[1].AddParagraph("Second section");

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-numbering-placement" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteSettingsWithoutReferencesRemainExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("No note references");
        source.AddFootnoteProperties(WordNumberFormat.UpperRoman);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-numbering-placement" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordSeparatorReferenceWithoutNotesRemainsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("No note references");
        source.OpenXmlDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!.AddChild(
            new W.FootnoteDocumentWideProperties(new W.FootnoteSpecialReference { Id = -2 }), true);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-separators" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void OdtSeparatorWithoutNotesRemainsExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!
            .Add(new XElement(style + "footnote-sep"));
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-configuration" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void InheritedWordReferenceStyleFormattingRemainsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Style baseStyle = source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(candidate => candidate.StyleId?.Value == "DefaultParagraphFont");
        baseStyle.Append(new W.StyleRunProperties(new W.Color { Val = "AA0000" }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void HyperlinkWrappedWordNoteReferenceReportsContextLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Run anchor = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.Run>().Single(run => run.GetFirstChild<W.FootnoteReference>() != null);
        anchor.Parent!.ReplaceChild(new W.Hyperlink((W.Run)anchor.CloneNode(true)) {
            Anchor = "Destination"
        }, anchor);

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-position" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CustomizedWordReferenceStyleIsReportedAsReferenceFormattingLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Body");
        W.Style style = source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(candidate => candidate.StyleId?.Value == "FootnoteReference");
        style.StyleRunProperties!.Append(new W.Color { Val = "AA0000" });

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting");
    }

    [Fact]
    public void WordNoteReferenceSharingARunWithBreakReportsOrderLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Note body");
        W.Run noteRun = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.Run>().Single(run => run.GetFirstChild<W.FootnoteReference>() != null);
        noteRun.Append(new W.Break());

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-reference-position" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void WordNoteNumberingAndPlacementSettingsAreExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddFootnoteProperties(WordNumberFormat.UpperRoman, WordFootnotePosition.BeneathText,
            WordNoteNumberRestart.EachSection, 5);
        source.AddEndnoteProperties(WordNumberFormat.LowerRoman, WordEndnotePosition.SectionEnd,
            WordNoteNumberRestart.EachSection, 3);
        source.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
        source.AddParagraph("Endnote anchor").AddEndNote("Endnote body");

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-numbering-placement" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 2);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void WordNoteBodyParagraphStylesAreExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Styled note body");
        var noteParagraphs = source.FootNotes[0].Paragraphs;
        Assert.NotNull(noteParagraphs);
        noteParagraphs![0].Style = WordParagraphStyles.Heading1;

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CustomizedBuiltInWordNoteStyleIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Note body");
        W.Style style = source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(candidate => candidate.StyleId?.Value == "FootnoteText");
        style.StyleRunProperties!.GetFirstChild<W.FontSize>()!.Val = "28";

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void InheritedNormalWordNoteFormattingIsExplicitLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Anchor").AddFootNote("Note body");
        W.Style normal = source.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<W.Style>().Single(candidate => candidate.StyleId?.Value == "Normal");
        normal.Append(new W.StyleRunProperties(new W.Bold()));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void HeaderAndFooterNoteBodyEditsSurviveSaveAndReopen() {
        OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Header").AddFootnote("Header note");
        source.PageLayout.Footer.AddParagraph("Footer").AddEndnote("Footer note");
        OdtDocument loaded = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        loaded.PageLayout.Header.Paragraphs[0].Notes[0].AddParagraph("Additional header detail");
        loaded.PageLayout.Footer.Paragraphs[0].Notes[0].Paragraphs[0].Text = "Edited footer note";

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal("Additional header detail", reopened.PageLayout.Header.Paragraphs[0].Notes[0].Paragraphs[1].Text);
        Assert.Equal("Edited footer note", reopened.PageLayout.Footer.Paragraphs[0].Notes[0].Paragraphs[0].Text);
        Assert.NotEqual(reopened.PageLayout.Header.Paragraphs[0].Notes[0].Id,
            reopened.PageLayout.Footer.Paragraphs[0].Notes[0].Id);
    }

    [Fact]
    public void ExistingWordNoteFixtureKeepsBothNoteKindsThroughOdt() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "DocumentWithFootNotes.docx");
        using WordDocument source = WordDocument.Load(path);
        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.True(conversion.Value.Validate().IsValid);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 3);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "endnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        OdtNote[] notes = reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes).ToArray();
        Assert.Equal(3, notes.Count(note => note.Kind == OdtNoteKind.Footnote));
        Assert.Equal(2, notes.Count(note => note.Kind == OdtNoteKind.Endnote));
        Assert.Contains(notes, note => note.Paragraphs.Any(paragraph => paragraph.Text.Contains("first footnote", StringComparison.Ordinal)));
        Assert.Contains(notes, note => note.Paragraphs.Any(paragraph => paragraph.Text.Contains("1st end note", StringComparison.Ordinal)));
    }

    [Fact]
    public void WordFootnotesAndEndnotesSurviveOdtAndDocxPackages() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Before ");
        paragraph.AddFootNote("Footnote text");
        paragraph.AddText(" between ");
        paragraph.AddEndNote("Endnote text");
        paragraph.AddText(" after");

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocumentResult();
        Assert.True(toOdt.Value.Validate().IsValid);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "endnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
        Assert.DoesNotContain(toOdt.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" ||
            mapping.Feature == "note-separators");
        Assert.DoesNotContain(toOdt.Report.Mappings, mapping => mapping.Feature == "source-footnotes" ||
            mapping.Feature == "source-endnotes");

        using var package = new MemoryStream(toOdt.Value.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        OdtParagraph odtParagraph = Assert.Single(reopened.ContentBlocks).Paragraph!;
        Assert.Equal("Before  between  after", odtParagraph.Text);
        Assert.Equal(new OdtNoteKind?[] { OdtNoteKind.Footnote, OdtNoteKind.Endnote },
            odtParagraph.Notes.Select(note => note.Kind).ToArray());
        Assert.Equal("Footnote text", odtParagraph.Notes[0].Paragraphs[0].Text);
        Assert.Equal("Endnote text", odtParagraph.Notes[1].Paragraphs[0].Text);
        Assert.Equal(new[] { OdtInlineNodeKind.Span, OdtInlineNodeKind.Note, OdtInlineNodeKind.Span,
            OdtInlineNodeKind.Note, OdtInlineNodeKind.Span }, odtParagraph.InlineNodes.Select(node => node.Kind));
        Assert.Single(reopened.Paragraphs);

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocumentResult();
        Assert.DoesNotContain(toWord.Report.Mappings, mapping => mapping.Feature == "source-text-notes");
        using WordDocument word = toWord.Value;
        Assert.Empty(word.ValidateDocument());
        using WordDocument saved = WordDocument.Load(new MemoryStream(word.ToBytes()));
        WordParagraphSnapshot snapshot = Assert.Single(saved.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());
        Assert.Equal("Footnote text", Assert.Single(snapshot.Runs, run => run.Footnote != null)
            .Footnote!.Paragraphs[0].Text);
        Assert.Equal("Endnote text", Assert.Single(snapshot.Runs, run => run.Endnote != null)
            .Endnote!.Paragraphs[0].Text);
    }

    [Fact]
    public void OdtNoteBodiesDoNotBecomeBodyParagraphsAndRichBodyLossIsExplicit() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Body");
        OdtNote note = paragraph.AddFootnote("First note paragraph");
        note.AddParagraph("Second note paragraph");
        paragraph.AddEndnote("End note");
        Assert.Single(source.Paragraphs);
        Assert.Single(source.ContentBlocks);

        using var package = new MemoryStream(source.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        Assert.Equal(2, Assert.Single(reopened.Paragraphs).Notes.Count);
        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CustomOdtCitationsAndUnmodeledNoteMediaRemainExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtNote note = source.AddParagraph("Cited text").AddFootnote("Source text");
        note.Paragraphs[0].AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        byte[] modified = OdfTestPackageRewriter.Rewrite(source.ToBytes(), (name, bytes) => {
            if (name != "content.xml") return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            xml.Descendants(text + "note-citation").Single().Value = "*";
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(modified));
        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-citations" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void NestedOdtNoteBodyImageIsUnsupportedLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Anchor").AddFootnote("Body").Paragraphs[0];
        body.AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        byte[] nested = OdfTestPackageRewriter.Rewrite(source.ToBytes(), (name, bytes) => {
            if (name != "content.xml") return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XElement frame = xml.Descendants(draw + "frame").Single();
            frame.ReplaceWith(new XElement(text + "span", frame));
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
        OdtDocument loaded = OdtDocument.Load(new MemoryStream(nested));

        OdfConversionResult<WordDocument> conversion = loaded.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => loaded.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }));
    }

    [Fact]
    public void SharedWordFootnoteReferenceAcrossRunsIsUnsupportedLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("First").AddFootNote("Shared note");
        source.AddParagraph("Second");
        long? id = source.OpenXmlDocument.MainDocumentPart!.Document!.Body!
            .Descendants<W.FootnoteReference>().Single().Id?.Value;
        Assert.NotNull(id);
        source.OpenXmlDocument.MainDocumentPart.Document.Body.Elements<W.Paragraph>().Last()
            .Append(new W.Run(new W.FootnoteReference { Id = id.Value }));

        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }));
    }
}
