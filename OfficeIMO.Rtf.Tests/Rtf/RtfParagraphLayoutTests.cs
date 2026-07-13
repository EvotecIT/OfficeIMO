using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfParagraphLayoutTests {
    [Fact]
    public void Read_Binds_Paragraph_Layout_Controls_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi\pard\pagebb\keepn\keep\noline\hyphpar0\contextualspace\adjustright\nosnaplinegrid\nowidctlpar\outlinelevel2\sb120\sa240\sbauto1\saauto0\sl360\slmult0 Layout\par\pard Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(2, result.Document.Paragraphs.Count);
        RtfParagraph layout = result.Document.Paragraphs[0];
        Assert.Equal("Layout", layout.ToPlainText());
        Assert.True(layout.PageBreakBefore);
        Assert.True(layout.KeepWithNext);
        Assert.True(layout.KeepLinesTogether);
        Assert.True(layout.SuppressLineNumbers);
        Assert.False(layout.AutoHyphenation);
        Assert.True(layout.ContextualSpacing);
        Assert.True(layout.AdjustRightIndent);
        Assert.False(layout.SnapToLineGrid);
        Assert.False(layout.WidowControl);
        Assert.Equal(2, layout.OutlineLevel);
        Assert.Equal(120, layout.SpaceBeforeTwips);
        Assert.Equal(240, layout.SpaceAfterTwips);
        Assert.True(layout.SpaceBeforeAuto);
        Assert.False(layout.SpaceAfterAuto);
        Assert.Equal(360, layout.LineSpacingTwips);
        Assert.False(layout.LineSpacingMultiple);

        RtfParagraph plain = result.Document.Paragraphs[1];
        Assert.Equal("Plain", plain.ToPlainText());
        Assert.False(plain.PageBreakBefore);
        Assert.False(plain.KeepWithNext);
        Assert.False(plain.KeepLinesTogether);
        Assert.False(plain.SuppressLineNumbers);
        Assert.Null(plain.AutoHyphenation);
        Assert.Null(plain.ContextualSpacing);
        Assert.Null(plain.AdjustRightIndent);
        Assert.Null(plain.SnapToLineGrid);
        Assert.Null(plain.WidowControl);
        Assert.Null(plain.OutlineLevel);
        Assert.Null(plain.SpaceBeforeTwips);
        Assert.Null(plain.SpaceAfterTwips);
        Assert.Null(plain.SpaceBeforeAuto);
        Assert.Null(plain.SpaceAfterAuto);
        Assert.Null(plain.LineSpacingTwips);
        Assert.Null(plain.LineSpacingMultiple);
    }

    [Fact]
    public void Write_And_Read_Paragraph_Layout_Controls_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Layout")
            .SetPagination(pageBreakBefore: true, keepWithNext: true, keepLinesTogether: true, widowControl: false, suppressLineNumbers: true, autoHyphenation: false)
            .SetContextualSpacing()
            .SetAdjustRightIndent()
            .SetSnapToLineGrid(false)
            .SetOutlineLevel(2)
            .SetParagraphSpacing(beforeTwips: 120, afterTwips: 240, beforeAuto: true, afterAuto: false)
            .SetLineSpacing(360, multiple: false);
        document.AddParagraph("Hyphenate")
            .SetPagination(autoHyphenation: true);
        document.AddParagraph("No contextual")
            .SetContextualSpacing(false);
        document.AddParagraph("No adjust")
            .SetAdjustRightIndent(false);
        document.AddParagraph("Snap")
            .SetSnapToLineGrid();
        document.AddParagraph("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\pagebb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\keepn", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\keep", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\noline", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\hyphpar0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\contextualspace\adjustright\nosnaplinegrid\nowidctlpar", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\hyphpar\ql Hyphenate", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\contextualspace0\ql No contextual", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\adjustright0\ql No adjust", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nosnaplinegrid0\ql Snap", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nowidctlpar", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\outlinelevel2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sb120", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sa240", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sbauto1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\saauto0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sl360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\slmult0", rtf, StringComparison.Ordinal);

        RtfParagraph layout = read.Document.Paragraphs[0];
        Assert.True(layout.PageBreakBefore);
        Assert.True(layout.KeepWithNext);
        Assert.True(layout.KeepLinesTogether);
        Assert.True(layout.SuppressLineNumbers);
        Assert.False(layout.AutoHyphenation);
        Assert.True(layout.ContextualSpacing);
        Assert.True(layout.AdjustRightIndent);
        Assert.False(layout.SnapToLineGrid);
        Assert.False(layout.WidowControl);
        Assert.Equal(2, layout.OutlineLevel);
        Assert.Equal(120, layout.SpaceBeforeTwips);
        Assert.Equal(240, layout.SpaceAfterTwips);
        Assert.True(layout.SpaceBeforeAuto);
        Assert.False(layout.SpaceAfterAuto);
        Assert.Equal(360, layout.LineSpacingTwips);
        Assert.False(layout.LineSpacingMultiple);

        RtfParagraph hyphenate = read.Document.Paragraphs[1];
        Assert.True(hyphenate.AutoHyphenation);
        Assert.Null(hyphenate.ContextualSpacing);
        Assert.Null(hyphenate.AdjustRightIndent);
        Assert.Null(hyphenate.SnapToLineGrid);

        RtfParagraph noContextual = read.Document.Paragraphs[2];
        Assert.False(noContextual.ContextualSpacing);

        RtfParagraph noAdjust = read.Document.Paragraphs[3];
        Assert.False(noAdjust.AdjustRightIndent);

        RtfParagraph snap = read.Document.Paragraphs[4];
        Assert.True(snap.SnapToLineGrid);

        RtfParagraph plain = read.Document.Paragraphs[5];
        Assert.False(plain.PageBreakBefore);
        Assert.False(plain.KeepWithNext);
        Assert.False(plain.KeepLinesTogether);
        Assert.False(plain.SuppressLineNumbers);
        Assert.Null(plain.AutoHyphenation);
        Assert.Null(plain.ContextualSpacing);
        Assert.Null(plain.AdjustRightIndent);
        Assert.Null(plain.SnapToLineGrid);
        Assert.Null(plain.WidowControl);
        Assert.Null(plain.OutlineLevel);
        Assert.Null(plain.SpaceBeforeTwips);
        Assert.Null(plain.SpaceAfterTwips);
        Assert.Null(plain.SpaceBeforeAuto);
        Assert.Null(plain.SpaceAfterAuto);
        Assert.Null(plain.LineSpacingTwips);
        Assert.Null(plain.LineSpacingMultiple);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Paragraph_Layout_Controls() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Layout");
        paragraph.PageBreakBefore = true;
        paragraph.KeepWithNext = true;
        paragraph.KeepLinesTogether = true;
        paragraph.LineSpacingBefore = 120;
        paragraph.LineSpacingAfter = 240;
        paragraph.LineSpacing = 360;
        paragraph.LineSpacingRule = LineSpacingRuleValues.Exact;
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines ??= new SpacingBetweenLines();
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines.BeforeAutoSpacing = true;
        paragraph._paragraph.ParagraphProperties.SpacingBetweenLines.AfterAutoSpacing = false;
        paragraph._paragraph.ParagraphProperties.SuppressLineNumbers = new SuppressLineNumbers();
        paragraph._paragraph.ParagraphProperties.SuppressAutoHyphens = new SuppressAutoHyphens();
        paragraph._paragraph.ParagraphProperties.ContextualSpacing = new ContextualSpacing();
        paragraph._paragraph.ParagraphProperties.AdjustRightIndent = new AdjustRightIndent();
        paragraph._paragraph.ParagraphProperties.SnapToGrid = new SnapToGrid { Val = false };
        paragraph._paragraph.ParagraphProperties.WidowControl = new WidowControl { Val = false };
        paragraph._paragraph.ParagraphProperties.OutlineLevel = new OutlineLevel { Val = 2 };

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.True(rtfParagraph.PageBreakBefore);
        Assert.True(rtfParagraph.KeepWithNext);
        Assert.True(rtfParagraph.KeepLinesTogether);
        Assert.True(rtfParagraph.SuppressLineNumbers);
        Assert.False(rtfParagraph.AutoHyphenation);
        Assert.True(rtfParagraph.ContextualSpacing);
        Assert.True(rtfParagraph.AdjustRightIndent);
        Assert.False(rtfParagraph.SnapToLineGrid);
        Assert.False(rtfParagraph.WidowControl);
        Assert.Equal(2, rtfParagraph.OutlineLevel);
        Assert.Equal(120, rtfParagraph.SpaceBeforeTwips);
        Assert.Equal(240, rtfParagraph.SpaceAfterTwips);
        Assert.True(rtfParagraph.SpaceBeforeAuto);
        Assert.False(rtfParagraph.SpaceAfterAuto);
        Assert.Equal(360, rtfParagraph.LineSpacingTwips);
        Assert.False(rtfParagraph.LineSpacingMultiple);
        Assert.Contains(@"\pagebb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\keepn", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\keep", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\noline", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\hyphpar0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\contextualspace", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\adjustright", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nosnaplinegrid", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nowidctlpar", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\outlinelevel2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sb120", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sa240", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sbauto1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\saauto0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sl360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\slmult0", rtf, StringComparison.Ordinal);

        WordParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.True(roundTripParagraph.PageBreakBefore);
        Assert.True(roundTripParagraph.KeepWithNext);
        Assert.True(roundTripParagraph.KeepLinesTogether);
        Assert.NotNull(roundTripParagraph._paragraphProperties?.SuppressLineNumbers);
        Assert.NotNull(roundTripParagraph._paragraphProperties?.SuppressAutoHyphens);
        Assert.NotNull(roundTripParagraph._paragraphProperties?.ContextualSpacing);
        Assert.NotNull(roundTripParagraph._paragraphProperties?.AdjustRightIndent);
        Assert.False(roundTripParagraph._paragraphProperties?.SnapToGrid?.Val?.Value ?? true);
        Assert.False(roundTripParagraph._paragraphProperties?.WidowControl?.Val?.Value ?? true);
        Assert.Equal(2, roundTripParagraph._paragraphProperties?.OutlineLevel?.Val?.Value);
        Assert.Equal(120, roundTripParagraph.LineSpacingBefore);
        Assert.Equal(240, roundTripParagraph.LineSpacingAfter);
        Assert.True(roundTripParagraph._paragraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value ?? false);
        Assert.False(roundTripParagraph._paragraphProperties?.SpacingBetweenLines?.AfterAutoSpacing?.Value ?? true);
        Assert.Equal(360, roundTripParagraph.LineSpacing);
        Assert.Equal(LineSpacingRuleValues.Exact, roundTripParagraph.LineSpacingRule);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Applies_Paragraph_Layout_Controls() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Layout")
            .SetPagination(pageBreakBefore: true, keepWithNext: true, keepLinesTogether: true, widowControl: true, suppressLineNumbers: true, autoHyphenation: false)
            .SetContextualSpacing()
            .SetAdjustRightIndent()
            .SetSnapToLineGrid(false)
            .SetOutlineLevel(3)
            .SetParagraphSpacing(beforeTwips: 180, afterTwips: 300, beforeAuto: false, afterAuto: true)
            .SetLineSpacing(480, multiple: true);

        using WordDocument word = document.ToWordDocument();

        WordParagraph paragraph = Assert.Single(word.Paragraphs);
        Assert.True(paragraph.PageBreakBefore);
        Assert.True(paragraph.KeepWithNext);
        Assert.True(paragraph.KeepLinesTogether);
        Assert.NotNull(paragraph._paragraphProperties?.SuppressLineNumbers);
        Assert.NotNull(paragraph._paragraphProperties?.SuppressAutoHyphens);
        Assert.NotNull(paragraph._paragraphProperties?.ContextualSpacing);
        Assert.NotNull(paragraph._paragraphProperties?.AdjustRightIndent);
        Assert.False(paragraph._paragraphProperties?.SnapToGrid?.Val?.Value ?? true);
        Assert.True(paragraph._paragraphProperties?.WidowControl?.Val?.Value ?? false);
        Assert.Equal(3, paragraph._paragraphProperties?.OutlineLevel?.Val?.Value);
        Assert.Equal(180, paragraph.LineSpacingBefore);
        Assert.Equal(300, paragraph.LineSpacingAfter);
        Assert.False(paragraph._paragraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value ?? true);
        Assert.True(paragraph._paragraphProperties?.SpacingBetweenLines?.AfterAutoSpacing?.Value ?? false);
        Assert.Equal(480, paragraph.LineSpacing);
        Assert.Equal(LineSpacingRuleValues.Auto, paragraph.LineSpacingRule);
    }

    [Fact]
    public void Read_Binds_Paragraph_Direction_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi\pard\rtlpar RTL\par\pard Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(2, result.Document.Paragraphs.Count);
        Assert.Equal(RtfTextDirection.RightToLeft, result.Document.Paragraphs[0].Direction);
        Assert.Null(result.Document.Paragraphs[1].Direction);
    }

    [Fact]
    public void Write_And_Read_Paragraph_Direction_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("RTL").SetDirection(RtfTextDirection.RightToLeft);
        document.AddParagraph("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\pard\rtlpar\ql RTL", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pard\ql Plain", rtf, StringComparison.Ordinal);
        Assert.Equal(RtfTextDirection.RightToLeft, result.Document.Paragraphs[0].Direction);
        Assert.Null(result.Document.Paragraphs[1].Direction);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Paragraph_Direction() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("RTL");
        paragraph.BiDi = true;

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal(RtfTextDirection.RightToLeft, rtfParagraph.Direction);
        Assert.Contains(@"\rtlpar", rtf, StringComparison.Ordinal);
        Assert.True(Assert.Single(roundTrip.Paragraphs).BiDi);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Paragraph_Direction() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("RTL").SetDirection(RtfTextDirection.RightToLeft);

        using WordDocument word = document.ToWordDocument();

        Assert.True(Assert.Single(word.Paragraphs).BiDi);
    }

    [Fact]
    public void Read_Binds_Paragraph_Frame_Positioning_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi\pard\absw5040\absh-720\phpg\posxc\pvpg\posyt\abslock\absnoovrlp0\nowrap\dxfrtext173\dfrmtxtx240\dfrmtxty360\overlay\dropcapli3\dropcapt2 Framed\par\pard Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(2, result.Document.Paragraphs.Count);
        RtfParagraph framed = result.Document.Paragraphs[0];
        Assert.Equal("Framed", framed.ToPlainText());
        Assert.True(framed.Frame.HasAnyValue);
        Assert.Equal(5040, framed.Frame.WidthTwips);
        Assert.Equal(-720, framed.Frame.HeightTwips);
        Assert.Equal(RtfParagraphFrameHorizontalAnchor.Page, framed.Frame.HorizontalAnchor);
        Assert.Equal(RtfParagraphFrameHorizontalPosition.Center, framed.Frame.HorizontalPosition);
        Assert.Null(framed.Frame.HorizontalPositionTwips);
        Assert.Equal(RtfParagraphFrameVerticalAnchor.Page, framed.Frame.VerticalAnchor);
        Assert.Equal(RtfParagraphFrameVerticalPosition.Top, framed.Frame.VerticalPosition);
        Assert.Null(framed.Frame.VerticalPositionTwips);
        Assert.True(framed.Frame.AnchorLocked);
        Assert.False(framed.Frame.NoOverlap);
        Assert.True(framed.Frame.NoWrap);
        Assert.Equal(173, framed.Frame.TextWrapDistanceTwips);
        Assert.Equal(240, framed.Frame.TextWrapDistanceHorizontalTwips);
        Assert.Equal(360, framed.Frame.TextWrapDistanceVerticalTwips);
        Assert.True(framed.Frame.OverlayText);
        Assert.Equal(3, framed.Frame.DropCapLines);
        Assert.Equal(RtfDropCapKind.Margin, framed.Frame.DropCapKind);

        RtfParagraph plain = result.Document.Paragraphs[1];
        Assert.Equal("Plain", plain.ToPlainText());
        Assert.False(plain.Frame.HasAnyValue);
    }

    [Fact]
    public void Write_And_Read_Paragraph_Frame_Positioning_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Framed").SetFrame(frame => {
            frame.SetSize(widthTwips: 3600, heightTwips: 0)
                .SetAnchors(RtfParagraphFrameHorizontalAnchor.Margin, RtfParagraphFrameVerticalAnchor.Paragraph)
                .SetPosition(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, horizontalTwips: -180, RtfParagraphFrameVerticalPosition.Absolute, verticalTwips: 720)
                .SetWrapping(noWrap: true, allDirectionsTwips: 120, horizontalTwips: 240, verticalTwips: 360, overlayText: true, noOverlap: true)
                .SetDropCap(2, RtfDropCapKind.InText);
            frame.AnchorLocked = true;
        });
        document.AddParagraph("Inside").SetFrame(frame => {
            frame.SetAnchors(RtfParagraphFrameHorizontalAnchor.Column, RtfParagraphFrameVerticalAnchor.Margin)
                .SetPosition(RtfParagraphFrameHorizontalPosition.Inside, null, RtfParagraphFrameVerticalPosition.Outside, null);
        });
        document.AddParagraph("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\absw3600\absh0\phmrg\posnegx-180\pvpara\posy720\abslock\absnoovrlp1\nowrap\dxfrtext120\dfrmtxtx240\dfrmtxty360\overlay\dropcapli2\dropcapt1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\phcol\posxi\pvmrg\posyout\ql Inside", rtf, StringComparison.Ordinal);
        RtfParagraph framed = read.Document.Paragraphs[0];
        Assert.Equal(3600, framed.Frame.WidthTwips);
        Assert.Equal(0, framed.Frame.HeightTwips);
        Assert.Equal(RtfParagraphFrameHorizontalAnchor.Margin, framed.Frame.HorizontalAnchor);
        Assert.Equal(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, framed.Frame.HorizontalPosition);
        Assert.Equal(-180, framed.Frame.HorizontalPositionTwips);
        Assert.Equal(RtfParagraphFrameVerticalAnchor.Paragraph, framed.Frame.VerticalAnchor);
        Assert.Equal(RtfParagraphFrameVerticalPosition.Absolute, framed.Frame.VerticalPosition);
        Assert.Equal(720, framed.Frame.VerticalPositionTwips);
        Assert.True(framed.Frame.AnchorLocked);
        Assert.True(framed.Frame.NoOverlap);
        Assert.True(framed.Frame.NoWrap);
        Assert.Equal(120, framed.Frame.TextWrapDistanceTwips);
        Assert.Equal(240, framed.Frame.TextWrapDistanceHorizontalTwips);
        Assert.Equal(360, framed.Frame.TextWrapDistanceVerticalTwips);
        Assert.True(framed.Frame.OverlayText);
        Assert.Equal(2, framed.Frame.DropCapLines);
        Assert.Equal(RtfDropCapKind.InText, framed.Frame.DropCapKind);

        RtfParagraph inside = read.Document.Paragraphs[1];
        Assert.Equal(RtfParagraphFrameHorizontalAnchor.Column, inside.Frame.HorizontalAnchor);
        Assert.Equal(RtfParagraphFrameHorizontalPosition.Inside, inside.Frame.HorizontalPosition);
        Assert.Equal(RtfParagraphFrameVerticalAnchor.Margin, inside.Frame.VerticalAnchor);
        Assert.Equal(RtfParagraphFrameVerticalPosition.Outside, inside.Frame.VerticalPosition);

        Assert.False(read.Document.Paragraphs[2].Frame.HasAnyValue);
    }
}
