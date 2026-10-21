using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private const string ZeroNotePrefix = "ZEROREFPREFIX";
    private const string ZeroNoteSuffix = "ZEROREFSUFFIX";
    private const string ZeroNoteBody = "ZERONOTEBODY";
    private const string ZeroNoteFollowingBody = "AFTERZERONOTE";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NoteParagraphs_OrdinaryZeroIdentifierRetainsPublicAccessAfterDocxReload(bool endnote, bool explicitNormal) {
        using WordDocument source = CreateOrdinaryZeroNoteDocument(endnote, explicitNormal);
        AssertOrdinaryZeroNotePublicAccess(source, endnote);

        using WordDocument reloaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(reloaded.ValidateDocument());
        AssertOrdinaryZeroNotePublicAccess(reloaded, endnote);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_OrdinaryZeroIdentifierRetainsInlineReferenceAndNoteWithoutChangingSource(
        bool endnote, bool explicitNormal) {
        string path = Path.Combine(_directoryWithFiles, "OrdinaryZeroNote.docx");
        using (WordDocument source = CreateOrdinaryZeroNoteDocument(endnote, explicitNormal)) source.Save(path);
        string hashBefore = OrdinaryZeroNoteSourceHash(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.Empty(document.ValidateDocument());
        string[] xmlBefore = OrdinaryZeroNoteSourceXml(document, endnote);

        // This path uses the package and PDF directly. A failing public note
        // Paragraphs getter must not prevent observing a conversion omission.
        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        byte[] bytes = result.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in new[] { ZeroNotePrefix, ZeroNoteSuffix, ZeroNoteBody, ZeroNoteFollowingBody })
            Assert.Single(Regex.Matches(visible, Regex.Escape(token)));
        AssertZeroNoteInlineReferencePosition(pdf);
        Assert.Contains(ZeroNotePrefix + "1" + ZeroNoteSuffix, visible, StringComparison.Ordinal);
        if (endnote) {
            AssertZeroEndnoteFlowBodyAssociation(pdf);
            PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
            PdfLinkAnnotation annotation = Assert.Single(navigation.LinkAnnotations,
                link => link.DestinationName == "OfficeIMO.Endnote.0");
            Assert.True(annotation.IsNamedDestinationLink);
            Assert.True(annotation.Width > 0D && annotation.Height > 0D);
            Assert.Equal(Assert.Single(pdf.GetPages(), page => page.Text.Contains(ZeroNotePrefix, StringComparison.Ordinal)).Number,
                annotation.PageNumber);
            PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations,
                item => item.Name == "OfficeIMO.Endnote.0");
            Assert.Equal(Assert.Single(pdf.GetPages(), page => page.Text.Contains(ZeroNoteBody, StringComparison.Ordinal)).Number,
                destination.PageNumber);
        }
        else AssertPdfFootnoteAssociation(pdf, ZeroNotePrefix, ZeroNoteBody, "1");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");

        Assert.Equal(xmlBefore, OrdinaryZeroNoteSourceXml(document, endnote));
        Assert.Equal(hashBefore, OrdinaryZeroNoteSourceHash(path));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_PositiveNoteIdentifierKeepsSourcePositionInGenericBodyAndTableFlows(bool endnote, bool tableCell) {
        using WordDocument document = CreateOrdinaryZeroNoteDocument(endnote, explicitNormal: false);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        if (endnote) {
            Assert.Single(main.EndnotesPart!.Endnotes!.Elements<W.Endnote>()).Id = 1L;
            Assert.Single(main.Document.Descendants<W.EndnoteReference>()).Id = 1L;
        } else {
            Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>()).Id = 1L;
            Assert.Single(main.Document.Descendants<W.FootnoteReference>()).Id = 1L;
            // Multiple columns select the documented flow fallback rather than
            // page-footnote definitions; its inline marker still owns a position.
            document.Sections[0].ColumnCount = 2;
        }
        if (tableCell) {
            W.Paragraph owner = main.Document.Body!.Elements<W.Paragraph>().First();
            WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
            W.Paragraph placeholder = table.Rows[0].Cells[0].Paragraphs[0]._paragraph;
            W.Table sourceTable = placeholder.Ancestors<W.Table>().First();
            sourceTable.Remove();
            owner.InsertBeforeSelf(sourceTable);
            owner.Remove();
            placeholder.InsertBeforeSelf(owner);
            placeholder.Remove();
        }
        Assert.Empty(document.ValidateDocument());
        string[] xmlBefore = OrdinaryZeroNoteSourceXml(document, endnote);

        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in new[] { ZeroNotePrefix, ZeroNoteSuffix, ZeroNoteBody, ZeroNoteFollowingBody })
            Assert.Single(Regex.Matches(visible, Regex.Escape(token)));
        AssertZeroNoteInlineReferencePosition(pdf);
        Assert.Contains(ZeroNotePrefix + "1" + ZeroNoteSuffix, visible, StringComparison.Ordinal);
        if (endnote) Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeFootnotePlacementApproximated");
        else {
            var warning = Assert.Single(result.Warnings, item => item.Code == "NativeFootnotePlacementApproximated");
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        }
        Assert.Equal(xmlBefore, OrdinaryZeroNoteSourceXml(document, endnote));
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_InlineEquationAndEndnoteRetainSourcePositionInBodyAndTable(bool tableCell) {
        using WordDocument document = CreateFootnoteDiagnosticDocument();
        document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal);
        WordParagraph owner = tableCell
            ? document.AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0].SetText(ZeroNotePrefix)
            : document.AddParagraph(ZeroNotePrefix);
        owner.AddEquation("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>MATH</m:t></m:r></m:oMath>")
            .AddEndNote(ZeroNoteBody).AddText(ZeroNoteSuffix);
        document.AddParagraph(ZeroNoteFollowingBody);
        Assert.Empty(document.ValidateDocument());
        W.EndnoteReference allocatedReference = Assert.Single(owner._paragraph.Descendants<W.EndnoteReference>());
        Assert.True(allocatedReference.Id?.Value > 0);
        long sourceId = allocatedReference.Id!.Value;
        Assert.Single(document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<W.Endnote>(),
            note => note.Id?.Value == sourceId);
        string[] sourceOrder = owner._paragraph.Descendants().Where(element => element is W.EndnoteReference ||
            (element is W.Text || element is DocumentFormat.OpenXml.Math.Text) && element.InnerText.Length > 0)
            .Select(element => element is W.EndnoteReference sourceReference ? "reference:" + sourceReference.Id!.Value : element.InnerText).ToArray();
        Assert.Equal(new[] { ZeroNotePrefix, "MATH", "reference:" + sourceId, ZeroNoteSuffix }, sourceOrder);
        string[] xmlBefore = OrdinaryZeroNoteSourceXml(document, endnote: true);

        var result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string token in new[] { ZeroNotePrefix, "MATH", ZeroNoteSuffix, ZeroNoteBody, ZeroNoteFollowingBody })
            Assert.Single(Regex.Matches(visible, Regex.Escape(token)));
        Letter reference = AssertZeroNoteInlineReferencePosition(pdf);
        Page page = Assert.Single(pdf.GetPages(), item => item.Text.Contains("MATH", StringComparison.Ordinal));
        Letter[] equation = AnchoredRunningLetters(page, "MATH");
        Letter[] prefix = AnchoredRunningLetters(page, ZeroNotePrefix);
        Assert.True(equation[0].StartBaseLine.X >= prefix.Last().EndBaseLine.X - .15D);
        Assert.True(equation.Last().EndBaseLine.X <= reference.StartBaseLine.X + .15D);
        Assert.Contains(ZeroNotePrefix + "MATH1" + ZeroNoteSuffix, visible, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeBodyEquationUnsupported");
        Assert.Equal(xmlBefore, OrdinaryZeroNoteSourceXml(document, endnote: true));
        Assert.Empty(document.ValidateDocument());
    }

    private static void AssertZeroEndnoteFlowBodyAssociation(PdfPigDocument pdf) {
        Page page = Assert.Single(pdf.GetPages(), item => item.Text.Contains(ZeroNoteBody, StringComparison.Ordinal));
        Letter[] body = AnchoredRunningLetters(page, ZeroNoteBody);
        // The source authors a marker directly before the supplied note text;
        // association must not require a space invented by flattened output.
        Letter label = Assert.Single(page.Letters, letter => letter.Value == "1" &&
            letter.StartBaseLine.X < body[0].StartBaseLine.X &&
            letter.StartBaseLine.Y >= body[0].StartBaseLine.Y &&
            letter.StartBaseLine.Y - body[0].StartBaseLine.Y < body[0].PointSize);
        Assert.True(label.EndBaseLine.X <= body[0].StartBaseLine.X);
        Assert.True(label.PointSize < body[0].PointSize);
        Page following = Assert.Single(pdf.GetPages(), item => item.Text.Contains(ZeroNoteFollowingBody, StringComparison.Ordinal));
        Assert.True(following.Number <= page.Number);
        if (following.Number == page.Number)
            Assert.True(AnchoredRunningLetters(following, ZeroNoteFollowingBody)[0].StartBaseLine.Y > label.StartBaseLine.Y);
    }

    private static Letter AssertZeroNoteInlineReferencePosition(PdfPigDocument pdf) {
        var page = Assert.Single(pdf.GetPages(), item => item.Text.Contains(ZeroNotePrefix, StringComparison.Ordinal));
        var prefix = AnchoredRunningLetters(page, ZeroNotePrefix);
        var suffix = AnchoredRunningLetters(page, ZeroNoteSuffix);
        var reference = Assert.Single(page.Letters, letter => letter.Value == "1" &&
            letter.StartBaseLine.Y > prefix[0].StartBaseLine.Y &&
            letter.StartBaseLine.Y - prefix[0].StartBaseLine.Y < prefix[0].PointSize);
        Assert.Equal(prefix[0].StartBaseLine.Y, suffix[0].StartBaseLine.Y, 3);
        Assert.True(reference.PointSize < prefix[0].PointSize);
        Assert.True(reference.StartBaseLine.X < suffix[0].StartBaseLine.X,
            FormattableString.Invariant($"Inline reference X={reference.StartBaseLine.X:F6}, endX={reference.EndBaseLine.X:F6}; prefix endX={prefix.Last().EndBaseLine.X:F6}, suffix X={suffix[0].StartBaseLine.X:F6}."));
        // The existing reference-association budget covers PDF operator rounding.
        Assert.InRange(reference.StartBaseLine.X, prefix.Last().EndBaseLine.X - .15D, suffix[0].StartBaseLine.X + .15D);
        Assert.True(reference.EndBaseLine.X <= suffix[0].StartBaseLine.X + .15D);
        return reference;
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Remove_OrdinaryZeroIdentifierPreservesSharedRunAndOtherDefinitions(bool endnote, bool reload) {
        using WordDocument source = CreateOrdinaryZeroNoteDocument(endnote, explicitNormal: false);
        WordParagraph owner = source.Paragraphs.First();
        if (endnote) owner.AddText("ZEROSIBLINGTEXT").AddEndNote("KEPTNOTEBODY").AddFootNote("COMPANIONNOTEBODY");
        else owner.AddText("ZEROSIBLINGTEXT").AddFootNote("KEPTNOTEBODY").AddEndNote("COMPANIONNOTEBODY");
        var originalOwner = source._wordprocessingDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().First();
        var shared = new W.Run(new W.RunProperties(new W.Bold()));
        foreach (W.Run run in originalOwner.Elements<W.Run>().Skip(1).ToArray()) {
            foreach (var child in run.ChildElements.Where(child => child is not W.RunProperties).ToArray()) {
                child.Remove();
                shared.Append(child);
            }
            run.Remove();
        }
        originalOwner.Append(shared);
        Assert.Empty(source.ValidateDocument());
        using WordDocument? restored = reload ? WordDocument.Load(new MemoryStream(source.ToBytes())) : null;
        WordDocument document = restored ?? source;
        string retainedDefinition = Assert.Single(OrdinaryZeroNoteDefinitions(document, endnote), note => note.Id == 1).Xml;
        string companionDefinitions = OrdinaryZeroNoteSourceXml(document, !endnote)[1];
        string ownerText = document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().First().InnerText;

        if (endnote) Assert.Single(document.EndNotes, note => note.ReferenceId == 0).Remove();
        else Assert.Single(document.FootNotes, note => note.ReferenceId == 0).Remove();

        AssertZeroNoteRemovalSurvivors(document, endnote, retainedDefinition, companionDefinitions, ownerText);
        using WordDocument removed = WordDocument.Load(new MemoryStream(document.ToBytes()));
        AssertZeroNoteRemovalSurvivors(removed, endnote, retainedDefinition, companionDefinitions, ownerText);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Remove_MissingNoteIdentifierCannotAliasOrdinaryZeroOrDeleteSpecialStories(bool endnote) {
        using WordDocument document = CreateOrdinaryZeroNoteDocument(endnote, explicitNormal: false);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Settings settings = main.DocumentSettingsPart!.Settings!;
        WordParagraph owner = document.AddParagraph("MISSINGREFPREFIX");
        if (endnote) {
            main.EndnotesPart!.Endnotes!.PrependChild(new W.Endnote(
                new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) {
                Id = -2L, Type = W.FootnoteEndnoteValues.ContinuationSeparator
            });
            W.EndnoteDocumentWideProperties? properties = settings.GetFirstChild<W.EndnoteDocumentWideProperties>();
            if (properties == null) { properties = new W.EndnoteDocumentWideProperties(); settings.AddChild(properties, true); }
            properties.AddChild(new W.EndnoteSpecialReference { Id = -2 }, true);
            owner._paragraph.Append(new W.Run(new W.RunProperties(new W.Bold()),
                new W.EndnoteReference(), new W.Text("MISSINGREFSUFFIX")));
        } else {
            main.FootnotesPart!.Footnotes!.PrependChild(new W.Footnote(
                new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) {
                Id = -2L, Type = W.FootnoteEndnoteValues.ContinuationSeparator
            });
            W.FootnoteDocumentWideProperties? properties = settings.GetFirstChild<W.FootnoteDocumentWideProperties>();
            if (properties == null) { properties = new W.FootnoteDocumentWideProperties(); settings.AddChild(properties, true); }
            properties.AddChild(new W.FootnoteSpecialReference { Id = -2 }, true);
            owner._paragraph.Append(new W.Run(new W.RunProperties(new W.Bold()),
                new W.FootnoteReference(), new W.Text("MISSINGREFSUFFIX")));
        }
        // A reference without its required ID is deliberately malformed. It
        // must not resolve to zero or give removal authority over any story.
        Assert.NotEmpty(document.ValidateDocument());
        string definitionsBefore = OrdinaryZeroNoteSourceXml(document, endnote)[1];
        string settingsBefore = settings.OuterXml;
        var expectedOwner = (W.Paragraph)owner._paragraph.CloneNode(true);
        if (endnote) Assert.Single(expectedOwner.Descendants<W.EndnoteReference>()).Remove();
        else Assert.Single(expectedOwner.Descendants<W.FootnoteReference>()).Remove();

        if (endnote) {
            WordEndNote malformed = Assert.Single(document.EndNotes, note => note.ReferenceId == null);
            Assert.Null(malformed.Paragraphs);
            malformed.Remove();
        } else {
            WordFootNote malformed = Assert.Single(document.FootNotes, note => note.ReferenceId == null);
            Assert.Null(malformed.Paragraphs);
            malformed.Remove();
        }

        Assert.Equal(expectedOwner.OuterXml, owner._paragraph.OuterXml);
        Assert.Equal(definitionsBefore, OrdinaryZeroNoteSourceXml(document, endnote)[1]);
        Assert.Equal(settingsBefore, settings.OuterXml);
        AssertOrdinaryZeroNotePublicAccess(document, endnote);
        Assert.Empty(document.ValidateDocument());
    }

    private static void AssertZeroNoteRemovalSurvivors(WordDocument document, bool endnote,
        string retainedDefinition, string companionDefinitions, string ownerText) {
        Assert.Equal(new[] { retainedDefinition }, OrdinaryZeroNoteDefinitions(document, endnote).Select(note => note.Xml));
        Assert.Equal(companionDefinitions, OrdinaryZeroNoteSourceXml(document, !endnote)[1]);
        W.Paragraph owner = document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().First();
        Assert.Equal(ownerText, owner.InnerText);
        Assert.Equal(2, owner.Elements<W.Run>().Count());
        W.Run shared = owner.Elements<W.Run>().Last();
        Assert.NotNull(shared.RunProperties?.Bold);
        Assert.Equal(new long?[] { 1 }, endnote
            ? shared.Elements<W.EndnoteReference>().Select(reference => reference.Id?.Value)
            : shared.Elements<W.FootnoteReference>().Select(reference => reference.Id?.Value));
        Assert.Single(endnote ? shared.Elements<W.FootnoteReference>().Cast<object>() : shared.Elements<W.EndnoteReference>().Cast<object>());
        long? referenceId = endnote ? Assert.Single(document.EndNotes).ReferenceId : Assert.Single(document.FootNotes).ReferenceId;
        Assert.Equal(1L, referenceId);
        List<WordParagraph>? paragraphs = endnote ? Assert.Single(document.EndNotes).Paragraphs : Assert.Single(document.FootNotes).Paragraphs;
        Assert.NotNull(paragraphs);
        Assert.Equal("KEPTNOTEBODY", string.Concat(paragraphs!.Select(paragraph => paragraph.Text)));
        Assert.Empty(document.ValidateDocument());
    }

    private static (long? Id, string Xml)[] OrdinaryZeroNoteDefinitions(WordDocument document, bool endnote) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return endnote
            ? main.EndnotesPart!.Endnotes!.Elements<W.Endnote>().Select(note => (note.Id?.Value, note.OuterXml)).ToArray()
            : main.FootnotesPart!.Footnotes!.Elements<W.Footnote>().Select(note => (note.Id?.Value, note.OuterXml)).ToArray();
    }

    private static WordDocument CreateOrdinaryZeroNoteDocument(bool endnote, bool explicitNormal) {
        WordDocument document = CreateFootnoteDiagnosticDocument();
        WordParagraph paragraph = document.AddParagraph(ZeroNotePrefix);
        if (endnote) {
            document.Sections[0].AddEndnoteProperties(WordNumberFormat.Decimal);
            paragraph.AddEndNote(ZeroNoteBody).AddText(ZeroNoteSuffix);
        } else paragraph.AddFootNote(ZeroNoteBody).AddText(ZeroNoteSuffix);
        document.AddParagraph(ZeroNoteFollowingBody);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Settings settings = main.DocumentSettingsPart!.Settings!;
        // Type defines a special story. Remove those stories and their settings
        // references before assigning zero to this ordinary imported-note shape.
        if (endnote) {
            W.Endnotes notes = main.EndnotesPart!.Endnotes!;
            foreach (W.Endnote special in notes.Elements<W.Endnote>().Where(note =>
                note.Type != null && note.Type.Value != W.FootnoteEndnoteValues.Normal).ToArray()) special.Remove();
            settings.GetFirstChild<W.EndnoteDocumentWideProperties>()?.RemoveAllChildren<W.EndnoteSpecialReference>();
            W.Endnote ordinary = Assert.Single(notes.Elements<W.Endnote>());
            ordinary.Id = 0L;
            ordinary.Type = null;
            if (explicitNormal) ordinary.Type = W.FootnoteEndnoteValues.Normal;
            Assert.Single(main.Document.Descendants<W.EndnoteReference>()).Id = 0L;
        } else {
            W.Footnotes notes = main.FootnotesPart!.Footnotes!;
            foreach (W.Footnote special in notes.Elements<W.Footnote>().Where(note =>
                note.Type != null && note.Type.Value != W.FootnoteEndnoteValues.Normal).ToArray()) special.Remove();
            settings.GetFirstChild<W.FootnoteDocumentWideProperties>()?.RemoveAllChildren<W.FootnoteSpecialReference>();
            W.Footnote ordinary = Assert.Single(notes.Elements<W.Footnote>());
            ordinary.Id = 0L;
            ordinary.Type = null;
            if (explicitNormal) ordinary.Type = W.FootnoteEndnoteValues.Normal;
            Assert.Single(main.Document.Descendants<W.FootnoteReference>()).Id = 0L;
        }
        Assert.Empty(document.ValidateDocument());
        return document;
    }

    private static void AssertOrdinaryZeroNotePublicAccess(WordDocument document, bool endnote) {
        if (endnote) {
            WordEndNote[] handles = {
                Assert.Single(document.EndNotes),
                Assert.IsType<WordEndNote>(Assert.Single(document.Paragraphs, paragraph => paragraph.IsEndNote).EndNote)
            };
            foreach (WordEndNote note in handles) {
                Assert.Equal(0L, note.ReferenceId);
                Assert.NotNull(note.Paragraphs);
                Assert.Equal(ZeroNoteBody, string.Concat(note.Paragraphs!.Select(paragraph => paragraph.Text)));
            }
        } else {
            WordFootNote[] handles = {
                Assert.Single(document.FootNotes),
                Assert.IsType<WordFootNote>(Assert.Single(document.Paragraphs, paragraph => paragraph.IsFootNote).FootNote)
            };
            foreach (WordFootNote note in handles) {
                Assert.Equal(0L, note.ReferenceId);
                Assert.NotNull(note.Paragraphs);
                Assert.Equal(ZeroNoteBody, string.Concat(note.Paragraphs!.Select(paragraph => paragraph.Text)));
            }
        }
    }

    private static string[] OrdinaryZeroNoteSourceXml(WordDocument document, bool endnote) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        return new[] {
            main.Document.OuterXml,
            endnote ? main.EndnotesPart!.Endnotes!.OuterXml : main.FootnotesPart!.Footnotes!.OuterXml,
            main.StyleDefinitionsPart!.Styles!.OuterXml,
            main.DocumentSettingsPart!.Settings!.OuterXml
        };
    }

    private static string OrdinaryZeroNoteSourceHash(string path) {
        using SHA256 algorithm = SHA256.Create();
        return BitConverter.ToString(algorithm.ComputeHash(File.ReadAllBytes(path))).Replace("-", string.Empty);
    }
}
