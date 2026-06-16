using System;
using System.Linq;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_Metadata_Styles_Hyperlinks_And_List_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "RTF Contract";
        document.Info.Author = "OfficeIMO";
        document.AddStyle(1, "Heading 1");
        document.AddStyle(2, "Hyperlink", RtfStyleKind.Character).Additive = true;

        RtfParagraph heading = document.AddParagraph("Heading");
        heading.SetStyle(1).SetAlignment(RtfTextAlignment.Center);
        RtfParagraph item = document.AddParagraph();
        item.SetList(kind: RtfListKind.Bullet).SetIndentation(leftTwips: 720, firstLineTwips: -360);
        item.AddText("Docs at ");
        item.AddText("OfficeIMO").SetStyle(2).SetHyperlink(new Uri("https://github.com/EvotecIT/OfficeIMO"));

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
        Assert.Contains(@"{\info{\title RTF Contract}{\author OfficeIMO}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\stylesheet{\s1 Heading 1;}{\*\cs2\additive Hyperlink;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\listtable", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\listoverridetable", rtf, StringComparison.Ordinal);
        Assert.Contains(@"HYPERLINK ""https://github.com/EvotecIT/OfficeIMO""", rtf, StringComparison.Ordinal);
        Assert.Equal("RTF Contract", read.Document.Info.Title);
        Assert.Equal("OfficeIMO", read.Document.Info.Author);
        Assert.Contains(read.Document.Styles, style => style.Id == 1 && style.Name == "Heading 1");
        Assert.Contains(read.Document.Styles, style => style.Id == 2 && style.Kind == RtfStyleKind.Character && style.Additive);
        Assert.Equal(RtfTextAlignment.Center, read.Document.Paragraphs[0].Alignment);
        Assert.Equal(1, read.Document.Paragraphs[0].StyleId);
        Assert.Single(read.Document.ListDefinitions);
        Assert.Single(read.Document.ListOverrides);
        Assert.Equal(read.Document.ListOverrides[0].ListId, read.Document.Paragraphs[1].ListDefinitionId);
        Assert.Equal(RtfListKind.Bullet, read.Document.Paragraphs[1].ListKind);
        RtfField hyperlink = Assert.Single(read.Document.Paragraphs[1].Inlines.OfType<RtfField>());
        Assert.Equal(new Uri("https://github.com/EvotecIT/OfficeIMO"), hyperlink.Hyperlink);
        Assert.Equal("OfficeIMO", hyperlink.ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Rich_Font_Table_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfFont defaultFont = document.Fonts.Single(font => font.Id == 0);
        defaultFont.Family = RtfFontFamily.Swiss;
        defaultFont.Charset = 0;
        defaultFont.Pitch = 2;
        defaultFont.CodePage = 1252;
        defaultFont.Bias = 0;
        defaultFont.Panose = "020F0502020204030204";
        defaultFont.AlternateName = "Arial";
        defaultFont.NonTaggedName = "Calibri";
        defaultFont.Embedding = new RtfFontEmbedding {
            Type = RtfEmbeddedFontType.TrueType,
            FileName = "Calibri.ttf",
            FileCodePage = 1252,
            Data = new byte[] { 1, 2, 3, 255 }
        };
        int monospaceFontId = document.AddFont("Consolas");
        document.Settings.SetDefaultFont(monospaceFontId);
        RtfFont monospace = document.Fonts.Single(font => font.Id == monospaceFontId);
        monospace.Family = RtfFontFamily.Modern;
        monospace.Charset = 238;
        monospace.Pitch = 1;
        monospace.CodePage = 1250;
        document.AddParagraph().AddText("Code").FontId = monospaceFontId;

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\deff1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\fonttbl{\f0\fswiss\fcharset0\fprq2\cpg1252\fbias0{\*\panose 020F0502020204030204}{\*\fname Calibri}{\*\fontemb\fttruetype{\*\fontfile\cpg1252 Calibri.ttf} 010203ff} Calibri{\*\falt Arial};}{\f1\fmodern\fcharset238\fprq1\cpg1250 Consolas;}}", rtf, StringComparison.Ordinal);
        Assert.Equal(monospaceFontId, read.Document.Settings.DefaultFontId);
        RtfFont readDefault = read.Document.Fonts.Single(font => font.Id == 0);
        Assert.Equal(RtfFontFamily.Swiss, readDefault.Family);
        Assert.Equal(0, readDefault.Charset);
        Assert.Equal(2, readDefault.Pitch);
        Assert.Equal(1252, readDefault.CodePage);
        Assert.Equal(0, readDefault.Bias);
        Assert.Equal("020F0502020204030204", readDefault.Panose);
        Assert.Equal("Arial", readDefault.AlternateName);
        Assert.Equal("Calibri", readDefault.NonTaggedName);
        Assert.NotNull(readDefault.Embedding);
        Assert.Equal(RtfEmbeddedFontType.TrueType, readDefault.Embedding.Type);
        Assert.Equal("Calibri.ttf", readDefault.Embedding.FileName);
        Assert.Equal(1252, readDefault.Embedding.FileCodePage);
        Assert.Equal(new byte[] { 1, 2, 3, 255 }, readDefault.Embedding.Data);
        RtfFont readMonospace = read.Document.Fonts.Single(font => font.Id == monospaceFontId);
        Assert.Equal(RtfFontFamily.Modern, readMonospace.Family);
        Assert.Equal(238, readMonospace.Charset);
        Assert.Equal(1, readMonospace.Pitch);
        Assert.Equal(1250, readMonospace.CodePage);
        Assert.Equal(monospaceFontId, Assert.Single(Assert.Single(read.Document.Paragraphs).Runs).FontId);
    }

    [Fact]
    public void Write_And_Read_Unicode_Font_And_Style_Names() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetUnicodeSkipCount(2);
        int fontId = document.AddFont("Zażółć Font");
        RtfStyle style = document.AddStyle(7, "Styl zażółć");
        RtfParagraph paragraph = document.AddParagraph("Styled");
        paragraph.StyleId = style.Id;
        paragraph.AddText(" text").FontId = fontId;

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\uc2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\u380??", rtf, StringComparison.Ordinal);
        Assert.Contains(read.Document.Fonts, font => font.Id == fontId && font.Name == "Zażółć Font");
        Assert.Contains(read.Document.Styles, readStyle => readStyle.Id == style.Id && readStyle.Name == "Styl zażółć");
        Assert.Equal(style.Id, Assert.Single(read.Document.Paragraphs).StyleId);
        Assert.Contains(Assert.Single(read.Document.Paragraphs).Runs, run => run.FontId == fontId);
    }

    [Fact]
    public void Write_And_Read_Rich_Color_Table_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int accentIndex = document.AddColor(68, 114, 196);
        RtfColor accent = document.Colors[accentIndex - 1];
        accent.ThemeColor = RtfThemeColor.AccentOne;
        accent.Tint = 40;
        int hyperlinkIndex = document.AddColor(5, 99, 193);
        RtfColor hyperlink = document.Colors[hyperlinkIndex - 1];
        hyperlink.ThemeColor = RtfThemeColor.Hyperlink;
        hyperlink.Shade = 25;
        document.AddParagraph().AddText("Accent").ForegroundColorIndex = accentIndex;

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\colortbl;\red68\green114\blue196\caccentone\ctint40;\red5\green99\blue193\chyperlink\cshade25;}", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.Colors,
            color => {
                Assert.Equal(RtfThemeColor.AccentOne, color.ThemeColor);
                Assert.Equal(40, color.Tint);
                Assert.Null(color.Shade);
            },
            color => {
                Assert.Equal(RtfThemeColor.Hyperlink, color.ThemeColor);
                Assert.Null(color.Tint);
                Assert.Equal(25, color.Shade);
            });
        Assert.Equal(accentIndex, Assert.Single(Assert.Single(read.Document.Paragraphs).Runs).ForegroundColorIndex);
    }

    [Fact]
    public void Write_And_Read_File_Table_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfFileReference local = document.AddFileReference(@"C:\Private\Resume\Edu\File2.docx", RtfFileSource.Ntfs);
        local.RelativePathStart = 18;
        RtfFileReference network = document.AddFileReference(@"\\Server\Share\Linked.docx", RtfFileSource.Ntfs | RtfFileSource.Network);
        network.OperatingSystemNumber = 42;
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\filetbl{\file\fid0\frelative18\fvalidntfs C:\\Private\\Resume\\Edu\\File2.docx}{\file\fid1\fosnum42\fvalidntfs\fnetwork \\\\Server\\Share\\Linked.docx}}", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.FileReferences,
            file => {
                Assert.Equal(0, file.Id);
                Assert.Equal(@"C:\Private\Resume\Edu\File2.docx", file.Path);
                Assert.Equal(18, file.RelativePathStart);
                Assert.Null(file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs, file.Sources);
            },
            file => {
                Assert.Equal(1, file.Id);
                Assert.Equal(@"\\Server\Share\Linked.docx", file.Path);
                Assert.Null(file.RelativePathStart);
                Assert.Equal(42, file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs | RtfFileSource.Network, file.Sources);
            });
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Xml_Namespace_Table() {
        RtfDocument document = RtfDocument.Create();
        document.AddXmlNamespace(2, "urn:contoso:custom");
        document.AddXmlNamespace(1, "http://schemas.example.test/word");
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\xmlnstbl{\xmlns1 http://schemas.example.test/word;}{\xmlns2 urn:contoso:custom;}}", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.XmlNamespaces,
            ns => {
                Assert.Equal(1, ns.Id);
                Assert.Equal("http://schemas.example.test/word", ns.Uri);
            },
            ns => {
                Assert.Equal(2, ns.Id);
                Assert.Equal("urn:contoso:custom", ns.Uri);
            });
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Rich_Stylesheet_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle heading = document.AddStyle(1, "Heading 1");
        heading.KeyCode = new RtfStyleKeyCode {
            Shift = true,
            Control = true,
            Key = "n"
        };
        heading.BasedOnStyleId = 0;
        heading.NextStyleId = 1;
        heading.LinkedStyleId = 2;
        heading.AutoUpdate = true;
        heading.Hidden = true;
        heading.Locked = true;
        heading.Personal = true;
        heading.Compose = true;
        heading.Reply = true;
        heading.SemiHidden = true;
        heading.UnhideWhenUsed = true;
        heading.QuickFormat = true;
        heading.Priority = 9;
        heading.RevisionSaveId = 123;
        RtfStyle character = document.AddStyle(2, "Character Link", RtfStyleKind.Character);
        character.Additive = true;
        character.LinkedStyleId = 1;
        document.AddParagraph("Heading").SetStyle(1);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\stylesheet{\s1{\*\keycode\shift\ctrl n}\sbasedon0\snext1\slink2\sautoupd\shidden\slocked\spersonal\scompose\sreply\ssemihidden\sunhideused\sqformat\spriority9\styrsid123 Heading 1;}{\*\cs2\slink1\additive Character Link;}}", rtf, StringComparison.Ordinal);
        RtfStyle readHeading = read.Document.Styles.Single(style => style.Id == 1 && style.Kind == RtfStyleKind.Paragraph);
        Assert.NotNull(readHeading.KeyCode);
        Assert.True(readHeading.KeyCode.Shift);
        Assert.True(readHeading.KeyCode.Control);
        Assert.False(readHeading.KeyCode.Alt);
        Assert.Null(readHeading.KeyCode.FunctionKey);
        Assert.Equal("n", readHeading.KeyCode.Key);
        Assert.Equal(0, readHeading.BasedOnStyleId);
        Assert.Equal(1, readHeading.NextStyleId);
        Assert.Equal(2, readHeading.LinkedStyleId);
        Assert.True(readHeading.AutoUpdate);
        Assert.True(readHeading.Hidden);
        Assert.True(readHeading.Locked);
        Assert.True(readHeading.Personal);
        Assert.True(readHeading.Compose);
        Assert.True(readHeading.Reply);
        Assert.True(readHeading.SemiHidden);
        Assert.True(readHeading.UnhideWhenUsed);
        Assert.True(readHeading.QuickFormat);
        Assert.Equal(9, readHeading.Priority);
        Assert.Equal(123, readHeading.RevisionSaveId);
        RtfStyle readCharacter = read.Document.Styles.Single(style => style.Id == 2 && style.Kind == RtfStyleKind.Character);
        Assert.True(readCharacter.Additive);
        Assert.Equal(1, readCharacter.LinkedStyleId);
    }

    [Fact]
    public void Write_And_Read_Stylesheet_Direct_Character_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int fontId = document.AddFont("Consolas");
        int red = document.AddColor(255, 0, 0);
        int yellow = document.AddColor(255, 255, 0);
        RtfStyle style = document.AddStyle(7, "Accent");
        style.Bold = true;
        style.Italic = true;
        style.UnderlineStyle = RtfUnderlineStyle.Double;
        style.FontSize = 15.5;
        style.FontId = fontId;
        style.ForegroundColorIndex = red;
        style.HighlightColorIndex = yellow;
        document.AddParagraph("Styled").SetStyle(style.Id);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\stylesheet{\s7\b\i\uldb\fs31\f1\cf1\highlight2 Accent;}}", rtf, StringComparison.Ordinal);
        RtfStyle readStyle = Assert.Single(read.Document.Styles);
        Assert.Equal(style.Id, readStyle.Id);
        Assert.Equal("Accent", readStyle.Name);
        Assert.Equal(true, readStyle.Bold);
        Assert.Equal(true, readStyle.Italic);
        Assert.Equal(RtfUnderlineStyle.Double, readStyle.UnderlineStyle);
        Assert.Equal(15.5, readStyle.FontSize);
        Assert.Equal(fontId, readStyle.FontId);
        Assert.Equal(red, readStyle.ForegroundColorIndex);
        Assert.Equal(yellow, readStyle.HighlightColorIndex);
        Assert.Equal(style.Id, Assert.Single(read.Document.Paragraphs).StyleId);
    }

    [Fact]
    public void Write_And_Read_Stylesheet_Direct_Paragraph_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        RtfStyle style = document.AddStyle(8, "Block");
        style.PageBreakBefore = true;
        style.KeepWithNext = true;
        style.KeepLinesTogether = true;
        style.SuppressLineNumbers = true;
        style.AutoHyphenation = true;
        style.ContextualSpacing = false;
        style.AdjustRightIndent = true;
        style.SnapToLineGrid = false;
        style.WidowControl = false;
        style.OutlineLevel = 3;
        style.ParagraphDirection = RtfTextDirection.LeftToRight;
        style.SetFrame(frame => frame
            .SetSize(widthTwips: 3600, heightTwips: 0)
            .SetAnchors(RtfParagraphFrameHorizontalAnchor.Margin, RtfParagraphFrameVerticalAnchor.Paragraph)
            .SetPosition(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, horizontalTwips: -180, RtfParagraphFrameVerticalPosition.Absolute, verticalTwips: 720)
            .SetWrapping(noWrap: true, allDirectionsTwips: 120, horizontalTwips: 240, verticalTwips: 360, overlayText: true, noOverlap: true)
            .SetDropCap(2, RtfDropCapKind.InText));
        style.Frame.AnchorLocked = true;
        style.AddTabStop(2880, RtfTabAlignment.Right, RtfTabLeader.Dots);
        style.LeftIndentTwips = 720;
        style.RightIndentTwips = 360;
        style.FirstLineIndentTwips = -180;
        style.SpaceBeforeTwips = 120;
        style.SpaceAfterTwips = 240;
        style.SpaceBeforeAuto = false;
        style.SpaceAfterAuto = true;
        style.LineSpacingTwips = 360;
        style.LineSpacingMultiple = true;
        style.BackgroundColorIndex = red;
        style.ShadingForegroundColorIndex = blue;
        style.ShadingPatternPercent = 5000;
        style.ShadingPattern = RtfShadingPattern.DarkHorizontal;
        style.SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 12, colorIndex: red)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 8, colorIndex: blue)
            .SetBorder(RtfParagraphBorderSide.Bottom, RtfParagraphBorderStyle.Dotted)
            .SetBorder(RtfParagraphBorderSide.Right, RtfParagraphBorderStyle.Dashed);
        style.ParagraphAlignment = RtfTextAlignment.Justify;
        document.AddParagraph("Styled").SetStyle(style.Id);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\stylesheet{\s8\pagebb\keepn\keep\noline\hyphpar\contextualspace0\adjustright\nosnaplinegrid\nowidctlpar\outlinelevel3\ltrpar\absw3600\absh0\phmrg\posnegx-180\pvpara\posy720\abslock\absnoovrlp1\nowrap\dxfrtext120\dfrmtxtx240\dfrmtxty360\overlay\dropcapli2\dropcapt1\tldot\tqr\tx2880\li720\ri360\fi-180\sb120\sa240\sbauto0\saauto1\sl360\slmult1\cbpat1\cfpat2\shading5000\bgdkhoriz\brdrt\brdrs\brdrw12\brdrcf1\brdrl\brdrdb\brdrw8\brdrcf2\brdrb\brdrdot\brdrr\brdrdash\qj Block;}}", rtf, StringComparison.Ordinal);
        RtfStyle readStyle = Assert.Single(read.Document.Styles);
        Assert.Equal(style.Id, readStyle.Id);
        Assert.Equal("Block", readStyle.Name);
        Assert.Equal(true, readStyle.PageBreakBefore);
        Assert.Equal(true, readStyle.KeepWithNext);
        Assert.Equal(true, readStyle.KeepLinesTogether);
        Assert.Equal(true, readStyle.SuppressLineNumbers);
        Assert.Equal(true, readStyle.AutoHyphenation);
        Assert.Equal(false, readStyle.ContextualSpacing);
        Assert.Equal(true, readStyle.AdjustRightIndent);
        Assert.Equal(false, readStyle.SnapToLineGrid);
        Assert.Equal(false, readStyle.WidowControl);
        Assert.Equal(3, readStyle.OutlineLevel);
        Assert.Equal(RtfTextDirection.LeftToRight, readStyle.ParagraphDirection);
        Assert.Equal(3600, readStyle.Frame.WidthTwips);
        Assert.Equal(0, readStyle.Frame.HeightTwips);
        Assert.Equal(RtfParagraphFrameHorizontalAnchor.Margin, readStyle.Frame.HorizontalAnchor);
        Assert.Equal(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, readStyle.Frame.HorizontalPosition);
        Assert.Equal(-180, readStyle.Frame.HorizontalPositionTwips);
        Assert.Equal(RtfParagraphFrameVerticalAnchor.Paragraph, readStyle.Frame.VerticalAnchor);
        Assert.Equal(RtfParagraphFrameVerticalPosition.Absolute, readStyle.Frame.VerticalPosition);
        Assert.Equal(720, readStyle.Frame.VerticalPositionTwips);
        Assert.True(readStyle.Frame.AnchorLocked);
        Assert.Equal(true, readStyle.Frame.NoOverlap);
        Assert.True(readStyle.Frame.NoWrap);
        Assert.Equal(120, readStyle.Frame.TextWrapDistanceTwips);
        Assert.Equal(240, readStyle.Frame.TextWrapDistanceHorizontalTwips);
        Assert.Equal(360, readStyle.Frame.TextWrapDistanceVerticalTwips);
        Assert.True(readStyle.Frame.OverlayText);
        Assert.Equal(2, readStyle.Frame.DropCapLines);
        Assert.Equal(RtfDropCapKind.InText, readStyle.Frame.DropCapKind);
        RtfTabStop tabStop = Assert.Single(readStyle.TabStops);
        Assert.Equal(2880, tabStop.PositionTwips);
        Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
        Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
        Assert.Equal(720, readStyle.LeftIndentTwips);
        Assert.Equal(360, readStyle.RightIndentTwips);
        Assert.Equal(-180, readStyle.FirstLineIndentTwips);
        Assert.Equal(120, readStyle.SpaceBeforeTwips);
        Assert.Equal(240, readStyle.SpaceAfterTwips);
        Assert.Equal(false, readStyle.SpaceBeforeAuto);
        Assert.Equal(true, readStyle.SpaceAfterAuto);
        Assert.Equal(360, readStyle.LineSpacingTwips);
        Assert.Equal(true, readStyle.LineSpacingMultiple);
        Assert.Equal(red, readStyle.BackgroundColorIndex);
        Assert.Equal(blue, readStyle.ShadingForegroundColorIndex);
        Assert.Equal(5000, readStyle.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, readStyle.ShadingPattern);
        Assert.Equal(RtfParagraphBorderStyle.Single, readStyle.TopBorder.Style);
        Assert.Equal(12, readStyle.TopBorder.Width);
        Assert.Equal(red, readStyle.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, readStyle.LeftBorder.Style);
        Assert.Equal(8, readStyle.LeftBorder.Width);
        Assert.Equal(blue, readStyle.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Dotted, readStyle.BottomBorder.Style);
        Assert.Equal(RtfParagraphBorderStyle.Dashed, readStyle.RightBorder.Style);
        Assert.Equal(RtfTextAlignment.Justify, readStyle.ParagraphAlignment);
        Assert.Equal(style.Id, Assert.Single(read.Document.Paragraphs).StyleId);
    }

    [Fact]
    public void Write_And_Read_Table_Stylesheet_Row_And_Cell_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        RtfStyle style = document.AddStyle(9, "Table Grid", RtfStyleKind.Table);
        RtfTableRow row = style.TableRowFormat;
        row.KeepTogether = true;
        row.KeepWithNext = true;
        row.SetAutoFit(false)
            .SetDirection(RtfTableRowDirection.RightToLeft)
            .SetCellGap(120)
            .SetLeftIndent(720)
            .SetAlignment(RtfTableAlignment.Center)
            .SetShading(red, blue, patternValue: 5, patternPercent: 6250, RtfShadingPattern.DarkHorizontal)
            .SetPadding(topTwips: 120, leftTwips: 180)
            .SetSpacing(topTwips: 20)
            .SetPositionAnchors(RtfTableHorizontalAnchor.Page, RtfTableVerticalAnchor.Paragraph)
            .SetPosition(RtfTableHorizontalPosition.Absolute, horizontalTwips: 1440, RtfTableVerticalPosition.Bottom)
            .SetTextWrapDistances(leftTwips: 80);
        row.HeightTwips = 360;
        row.PreferredWidthUnit = RtfTableWidthUnit.Twips;
        row.PreferredWidth = 5000;
        row.NoOverlap = true;
        row.TopBorder.Style = RtfTableCellBorderStyle.Single;
        row.TopBorder.Width = 12;
        row.TopBorder.ColorIndex = red;

        RtfTableCell cell = row.AddCell(2400);
        cell.SetPreferredWidth(2400, RtfTableWidthUnit.Twips)
            .SetNoWrap()
            .SetFitText()
            .SetShading(red, blue, patternPercent: 3750, RtfShadingPattern.DarkHorizontal)
            .SetTextFlow(RtfTableCellTextFlow.LeftToRightTopToBottomVertical)
            .SetPadding(topTwips: 60);
        cell.VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        cell.TopBorder.Style = RtfTableCellBorderStyle.Double;
        cell.TopBorder.Width = 8;
        cell.TopBorder.ColorIndex = blue;
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\stylesheet{\*\ts9\tsrowd\trkeep\trkeepfollow\trautofit0\rtlrow\trrh360\trgaph120\trleft720\trftsWidth3\trwWidth5000\trcbpat1\trcfpat2\trpat5\trshdng6250\trbgdkhor\trpaddt120\trpaddft3\trpaddl180\trpaddfl3\trspdt20\trspdft3\tabsnoovrlp\tphpg\tpvpara\tposx1440\tposyb\tdfrmtxtLeft80\trqc\trbrdrt\brdrs\brdrw12\brdrcf1\clftsWidth3\clwWidth2400\clNoWrap\clFitText\clcbpat1\clcfpat2\clshdng3750\clbgdkhor\clvertalc\cltxlrtbv\clbrdrt\brdrdb\brdrw8\brdrcf2\clpadt60\clpadft3\cellx2400 Table Grid;}}", rtf, StringComparison.Ordinal);
        RtfStyle readStyle = Assert.Single(read.Document.Styles);
        Assert.Equal(style.Id, readStyle.Id);
        Assert.Equal(RtfStyleKind.Table, readStyle.Kind);
        Assert.Equal("Table Grid", readStyle.Name);

        RtfTableRow readRow = readStyle.TableRowFormat;
        Assert.True(readRow.KeepTogether);
        Assert.True(readRow.KeepWithNext);
        Assert.Equal(false, readRow.AutoFit);
        Assert.Equal(RtfTableRowDirection.RightToLeft, readRow.Direction);
        Assert.Equal(360, readRow.HeightTwips);
        Assert.Equal(120, readRow.CellGapTwips);
        Assert.Equal(720, readRow.LeftIndentTwips);
        Assert.Equal(RtfTableWidthUnit.Twips, readRow.PreferredWidthUnit);
        Assert.Equal(5000, readRow.PreferredWidth);
        Assert.Equal(red, readRow.BackgroundColorIndex);
        Assert.Equal(blue, readRow.ShadingForegroundColorIndex);
        Assert.Equal(5, readRow.ShadingPatternValue);
        Assert.Equal(6250, readRow.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, readRow.ShadingPattern);
        Assert.Equal(120, readRow.PaddingTopTwips);
        Assert.Equal(180, readRow.PaddingLeftTwips);
        Assert.Equal(20, readRow.SpacingTopTwips);
        Assert.True(readRow.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Page, readRow.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Paragraph, readRow.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Absolute, readRow.HorizontalPosition);
        Assert.Equal(1440, readRow.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Bottom, readRow.VerticalPosition);
        Assert.Equal(80, readRow.TextWrapLeftTwips);
        Assert.Equal(RtfTableAlignment.Center, readRow.Alignment);
        Assert.Equal(RtfTableCellBorderStyle.Single, readRow.TopBorder.Style);
        Assert.Equal(12, readRow.TopBorder.Width);
        Assert.Equal(red, readRow.TopBorder.ColorIndex);

        RtfTableCell readCell = Assert.Single(readRow.Cells);
        Assert.Equal(2400, readCell.RightBoundaryTwips);
        Assert.Equal(RtfTableWidthUnit.Twips, readCell.PreferredWidthUnit);
        Assert.Equal(2400, readCell.PreferredWidth);
        Assert.True(readCell.NoWrap);
        Assert.True(readCell.FitText);
        Assert.Equal(red, readCell.BackgroundColorIndex);
        Assert.Equal(blue, readCell.ShadingForegroundColorIndex);
        Assert.Equal(3750, readCell.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, readCell.ShadingPattern);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, readCell.VerticalAlignment);
        Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottomVertical, readCell.TextFlow);
        Assert.Equal(RtfTableCellBorderStyle.Double, readCell.TopBorder.Style);
        Assert.Equal(8, readCell.TopBorder.Width);
        Assert.Equal(blue, readCell.TopBorder.ColorIndex);
        Assert.Equal(60, readCell.PaddingTopTwips);
    }

}
