using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public void Read_Binds_Rich_Font_Table_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0\froman\fcharset0\fprq2\cpg1252\fbias0{\*\panose 02020603050405020304}{\*\fname Times New Roman}{\*\fontemb\fttruetype{\*\fontfile\cpg1252 TimesNewRoman.ttf}010203ff} Times New Roman{\*\falt Times};}{\f1\fmodern\fcharset238\fprq1\cpg1250 Consolas;}{\f2\fbidi\fcharset178\fprq0\cpg1256 Arabic Typesetting;}}\pard\f1 Code\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.Fonts,
            font => {
                Assert.Equal(0, font.Id);
                Assert.Equal("Times New Roman", font.Name);
                Assert.Equal(RtfFontFamily.Roman, font.Family);
                Assert.Equal(0, font.Charset);
                Assert.Equal(2, font.Pitch);
                Assert.Equal(1252, font.CodePage);
                Assert.Equal(0, font.Bias);
                Assert.Equal("02020603050405020304", font.Panose);
                Assert.Equal("Times", font.AlternateName);
                Assert.Equal("Times New Roman", font.NonTaggedName);
                Assert.NotNull(font.Embedding);
                Assert.Equal(RtfEmbeddedFontType.TrueType, font.Embedding.Type);
                Assert.Equal("TimesNewRoman.ttf", font.Embedding.FileName);
                Assert.Equal(1252, font.Embedding.FileCodePage);
                Assert.Equal(new byte[] { 1, 2, 3, 255 }, font.Embedding.Data);
            },
            font => {
                Assert.Equal(1, font.Id);
                Assert.Equal("Consolas", font.Name);
                Assert.Equal(RtfFontFamily.Modern, font.Family);
                Assert.Equal(238, font.Charset);
                Assert.Equal(1, font.Pitch);
                Assert.Equal(1250, font.CodePage);
            },
            font => {
                Assert.Equal(2, font.Id);
                Assert.Equal("Arabic Typesetting", font.Name);
                Assert.Equal(RtfFontFamily.Bidirectional, font.Family);
                Assert.Equal(178, font.Charset);
                Assert.Equal(0, font.Pitch);
                Assert.Equal(1256, font.CodePage);
            });
    }

    [Fact]
    public void Read_Binds_Default_Font_Id_From_Header() {
        const string rtf = @"{\rtf1\ansi\deff1{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}\pard Default\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(1, result.Document.Settings.DefaultFontId);
        Assert.Equal("Consolas", result.Document.Fonts.Single(font => font.Id == result.Document.Settings.DefaultFontId).Name);
    }

    [Fact]
    public void Read_Binds_Rich_Color_Table_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red68\green114\blue196\caccentone\ctint40;\red237\green125\blue49\caccenttwo\cshade25;\red5\green99\blue193\chyperlink;}\pard\cf1 Accent\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.Colors,
            color => {
                Assert.Equal((byte)68, color.Red);
                Assert.Equal((byte)114, color.Green);
                Assert.Equal((byte)196, color.Blue);
                Assert.Equal(RtfThemeColor.AccentOne, color.ThemeColor);
                Assert.Equal(40, color.Tint);
                Assert.Null(color.Shade);
            },
            color => {
                Assert.Equal((byte)237, color.Red);
                Assert.Equal((byte)125, color.Green);
                Assert.Equal((byte)49, color.Blue);
                Assert.Equal(RtfThemeColor.AccentTwo, color.ThemeColor);
                Assert.Null(color.Tint);
                Assert.Equal(25, color.Shade);
            },
            color => {
                Assert.Equal((byte)5, color.Red);
                Assert.Equal((byte)99, color.Green);
                Assert.Equal((byte)193, color.Blue);
                Assert.Equal(RtfThemeColor.Hyperlink, color.ThemeColor);
                Assert.Null(color.Tint);
                Assert.Null(color.Shade);
            });
    }

    [Fact]
    public void Read_Binds_File_Table_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\filetbl{\file\fid0\frelative18\fvalidntfs C:\\Private\\Resume\\Edu\\File2.docx}{\file\fid1\fosnum42\fvalidmac\fnetwork MacHD:Docs:Linked.doc}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.FileReferences,
            file => {
                Assert.Equal(0, file.Id);
                Assert.Equal(@"C:\Private\Resume\Edu\File2.docx", file.Path);
                Assert.Equal(18, file.RelativePathStart);
                Assert.Null(file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs, file.Sources);
            },
            file => {
                Assert.Equal(1, file.Id);
                Assert.Equal("MacHD:Docs:Linked.doc", file.Path);
                Assert.Null(file.RelativePathStart);
                Assert.Equal(42, file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Mac | RtfFileSource.Network, file.Sources);
            });
    }

    [Fact]
    public void Read_Binds_Xml_Namespace_Table_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\xmlnstbl{\xmlns1 http://schemas.example.test/word;}{\xmlns2 urn:contoso:custom;}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.XmlNamespaces,
            ns => {
                Assert.Equal(1, ns.Id);
                Assert.Equal("http://schemas.example.test/word", ns.Uri);
            },
            ns => {
                Assert.Equal(2, ns.Id);
                Assert.Equal("urn:contoso:custom", ns.Uri);
            });
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Rich_Stylesheet_Metadata_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\stylesheet{\s1{\*\keycode\shift\ctrl n}\sbasedon0\snext1\slink2\sautoupd\shidden\slocked\spersonal\scompose\sreply\ssemihidden\sunhideused\sqformat\spriority9\styrsid123 Heading 1;}{\*\cs2\additive\slink1 Character Link;}}\pard\s1 Heading\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.Styles,
            style => {
                Assert.Equal(1, style.Id);
                Assert.Equal(RtfStyleKind.Paragraph, style.Kind);
                Assert.Equal("Heading 1", style.Name);
                Assert.NotNull(style.KeyCode);
                Assert.True(style.KeyCode.Shift);
                Assert.True(style.KeyCode.Control);
                Assert.False(style.KeyCode.Alt);
                Assert.Null(style.KeyCode.FunctionKey);
                Assert.Equal("n", style.KeyCode.Key);
                Assert.Equal(0, style.BasedOnStyleId);
                Assert.Equal(1, style.NextStyleId);
                Assert.Equal(2, style.LinkedStyleId);
                Assert.True(style.AutoUpdate);
                Assert.True(style.Hidden);
                Assert.True(style.Locked);
                Assert.True(style.Personal);
                Assert.True(style.Compose);
                Assert.True(style.Reply);
                Assert.True(style.SemiHidden);
                Assert.True(style.UnhideWhenUsed);
                Assert.True(style.QuickFormat);
                Assert.Equal(9, style.Priority);
                Assert.Equal(123, style.RevisionSaveId);
            },
            style => {
                Assert.Equal(2, style.Id);
                Assert.Equal(RtfStyleKind.Character, style.Kind);
                Assert.Equal("Character Link", style.Name);
                Assert.True(style.Additive);
                Assert.Equal(1, style.LinkedStyleId);
            });
    }

    [Fact]
    public void Read_Binds_Stylesheet_Direct_Character_Formatting_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}{\colortbl;\red255\green0\blue0;\red255\green255\blue0;}{\stylesheet{\s1\b\i\ul\fs28\f1\cf1\highlight2 Emphasis;}{\*\cs2\b0\i0\ulnone\cf0 Plain Link;}}\pard\s1 Heading\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.Styles,
            style => {
                Assert.Equal(1, style.Id);
                Assert.Equal("Emphasis", style.Name);
                Assert.Equal(true, style.Bold);
                Assert.Equal(true, style.Italic);
                Assert.Equal(RtfUnderlineStyle.Single, style.UnderlineStyle);
                Assert.Equal(14, style.FontSize);
                Assert.Equal(1, style.FontId);
                Assert.Equal(1, style.ForegroundColorIndex);
                Assert.Equal(2, style.HighlightColorIndex);
            },
            style => {
                Assert.Equal(2, style.Id);
                Assert.Equal(RtfStyleKind.Character, style.Kind);
                Assert.Equal("Plain Link", style.Name);
                Assert.Equal(false, style.Bold);
                Assert.Equal(false, style.Italic);
                Assert.Equal(RtfUnderlineStyle.None, style.UnderlineStyle);
                Assert.Equal(0, style.ForegroundColorIndex);
                Assert.Null(style.HighlightColorIndex);
            });
    }

    [Fact]
    public void Read_Binds_Stylesheet_Direct_Paragraph_Formatting_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}{\stylesheet{\s3\pagebb\keepn\keep\noline\hyphpar0\contextualspace\adjustright0\nosnaplinegrid0\widctlpar\outlinelevel2\rtlpar\absw5040\absh-720\phpg\posxc\pvpg\posyt\abslock\absnoovrlp0\nowrap\dxfrtext173\dfrmtxtx240\dfrmtxty360\overlay\dropcapli3\dropcapt2\tqr\tldot\tx2880\li720\ri360\fi-180\sb120\sa240\sbauto1\saauto0\sl360\slmult1\cbpat1\cfpat2\shading5000\bghoriz\brdrt\brdrs\brdrw12\brdrcf1\brdrl\brdrdb\brdrw8\brdrcf2\brdrb\brdrdot\brdrr\brdrdash\qc Paragraph Format;}}\pard\s3 Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfStyle style = Assert.Single(result.Document.Styles);
        Assert.Equal(3, style.Id);
        Assert.Equal("Paragraph Format", style.Name);
        Assert.Equal(true, style.PageBreakBefore);
        Assert.Equal(true, style.KeepWithNext);
        Assert.Equal(true, style.KeepLinesTogether);
        Assert.Equal(true, style.SuppressLineNumbers);
        Assert.Equal(false, style.AutoHyphenation);
        Assert.Equal(true, style.ContextualSpacing);
        Assert.Equal(false, style.AdjustRightIndent);
        Assert.Equal(true, style.SnapToLineGrid);
        Assert.Equal(true, style.WidowControl);
        Assert.Equal(2, style.OutlineLevel);
        Assert.Equal(RtfTextDirection.RightToLeft, style.ParagraphDirection);
        Assert.Equal(5040, style.Frame.WidthTwips);
        Assert.Equal(-720, style.Frame.HeightTwips);
        Assert.Equal(RtfParagraphFrameHorizontalAnchor.Page, style.Frame.HorizontalAnchor);
        Assert.Equal(RtfParagraphFrameHorizontalPosition.Center, style.Frame.HorizontalPosition);
        Assert.Equal(RtfParagraphFrameVerticalAnchor.Page, style.Frame.VerticalAnchor);
        Assert.Equal(RtfParagraphFrameVerticalPosition.Top, style.Frame.VerticalPosition);
        Assert.True(style.Frame.AnchorLocked);
        Assert.Equal(false, style.Frame.NoOverlap);
        Assert.True(style.Frame.NoWrap);
        Assert.Equal(173, style.Frame.TextWrapDistanceTwips);
        Assert.Equal(240, style.Frame.TextWrapDistanceHorizontalTwips);
        Assert.Equal(360, style.Frame.TextWrapDistanceVerticalTwips);
        Assert.True(style.Frame.OverlayText);
        Assert.Equal(3, style.Frame.DropCapLines);
        Assert.Equal(RtfDropCapKind.Margin, style.Frame.DropCapKind);
        RtfTabStop tabStop = Assert.Single(style.TabStops);
        Assert.Equal(2880, tabStop.PositionTwips);
        Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
        Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
        Assert.Equal(720, style.LeftIndentTwips);
        Assert.Equal(360, style.RightIndentTwips);
        Assert.Equal(-180, style.FirstLineIndentTwips);
        Assert.Equal(120, style.SpaceBeforeTwips);
        Assert.Equal(240, style.SpaceAfterTwips);
        Assert.Equal(true, style.SpaceBeforeAuto);
        Assert.Equal(false, style.SpaceAfterAuto);
        Assert.Equal(360, style.LineSpacingTwips);
        Assert.Equal(true, style.LineSpacingMultiple);
        Assert.Equal(1, style.BackgroundColorIndex);
        Assert.Equal(2, style.ShadingForegroundColorIndex);
        Assert.Equal(5000, style.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.Horizontal, style.ShadingPattern);
        Assert.Equal(RtfParagraphBorderStyle.Single, style.TopBorder.Style);
        Assert.Equal(12, style.TopBorder.Width);
        Assert.Equal(1, style.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, style.LeftBorder.Style);
        Assert.Equal(8, style.LeftBorder.Width);
        Assert.Equal(2, style.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Dotted, style.BottomBorder.Style);
        Assert.Equal(RtfParagraphBorderStyle.Dashed, style.RightBorder.Style);
        Assert.Equal(RtfTextAlignment.Center, style.ParagraphAlignment);
        Assert.Equal(3, Assert.Single(result.Document.Paragraphs).StyleId);
    }

    [Fact]
    public void Read_Binds_Table_Stylesheet_Row_And_Cell_Formatting_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}{\stylesheet{\*\ts9\tsrowd\trgaph120\trrh360\trautofit0\rtlrow\trleft720\trqc\trftsWidth3\trwWidth5000\trcbpat1\trcfpat2\trpat5\trshdng6250\trbgdkhor\trpaddt120\trpaddft3\trpaddl180\trpaddfl3\trspdt20\trspdft3\tabsnoovrlp\tphpg\tpvpara\tposx1440\tposyb\tdfrmtxtLeft80\trbrdrt\brdrs\brdrw12\brdrcf1\clftsWidth3\clwWidth2400\clcbpat1\clcfpat2\clshdng3750\clbgdkhor\clvertalc\cltxlrtbv\clNoWrap\clFitText\clbrdrt\brdrdb\brdrw8\brdrcf2\clpadt60\clpadft3\cellx2400 Table Grid;}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfStyle style = Assert.Single(result.Document.Styles);
        Assert.Equal(9, style.Id);
        Assert.Equal(RtfStyleKind.Table, style.Kind);
        Assert.Equal("Table Grid", style.Name);
        RtfTableRow row = style.TableRowFormat;
        Assert.Equal(120, row.CellGapTwips);
        Assert.Equal(360, row.HeightTwips);
        Assert.Equal(false, row.AutoFit);
        Assert.Equal(RtfTableRowDirection.RightToLeft, row.Direction);
        Assert.Equal(720, row.LeftIndentTwips);
        Assert.Equal(RtfTableAlignment.Center, row.Alignment);
        Assert.Equal(RtfTableWidthUnit.Twips, row.PreferredWidthUnit);
        Assert.Equal(5000, row.PreferredWidth);
        Assert.Equal(1, row.BackgroundColorIndex);
        Assert.Equal(2, row.ShadingForegroundColorIndex);
        Assert.Equal(5, row.ShadingPatternValue);
        Assert.Equal(6250, row.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, row.ShadingPattern);
        Assert.Equal(120, row.PaddingTopTwips);
        Assert.Equal(180, row.PaddingLeftTwips);
        Assert.Equal(20, row.SpacingTopTwips);
        Assert.True(row.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Page, row.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Paragraph, row.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Absolute, row.HorizontalPosition);
        Assert.Equal(1440, row.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Bottom, row.VerticalPosition);
        Assert.Equal(80, row.TextWrapLeftTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, row.TopBorder.Style);
        Assert.Equal(12, row.TopBorder.Width);
        Assert.Equal(1, row.TopBorder.ColorIndex);

        RtfTableCell cell = Assert.Single(row.Cells);
        Assert.Equal(2400, cell.RightBoundaryTwips);
        Assert.Equal(RtfTableWidthUnit.Twips, cell.PreferredWidthUnit);
        Assert.Equal(2400, cell.PreferredWidth);
        Assert.Equal(1, cell.BackgroundColorIndex);
        Assert.Equal(2, cell.ShadingForegroundColorIndex);
        Assert.Equal(3750, cell.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, cell.ShadingPattern);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, cell.VerticalAlignment);
        Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottomVertical, cell.TextFlow);
        Assert.True(cell.NoWrap);
        Assert.True(cell.FitText);
        Assert.Equal(RtfTableCellBorderStyle.Double, cell.TopBorder.Style);
        Assert.Equal(8, cell.TopBorder.Width);
        Assert.Equal(2, cell.TopBorder.ColorIndex);
        Assert.Equal(60, cell.PaddingTopTwips);
    }

    [Fact]
    public void Read_And_Write_Rich_Info_Metadata_Timestamps_And_Statistics() {
        const string rtf = @"{\rtf1\ansi{\info{\title Title}{\subject Subject}{\author Author}{\manager Manager}{\company Company}{\operator Operator}{\category Category}{\keywords one,two}{\doccomm Comment}{\hlinkbase https://example.test/}{\creatim\yr2026\mo6\dy15\hr10\min20\sec30}{\revtim\yr2026\mo6\dy16\hr11\min21\sec31}\edmins42\nofpages7\nofwords120\nofchars600\nofcharsws700\vern123}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfDocumentInfo info = result.Document.Info;

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("Title", info.Title);
        Assert.Equal("Subject", info.Subject);
        Assert.Equal("Author", info.Author);
        Assert.Equal("Manager", info.Manager);
        Assert.Equal("Company", info.Company);
        Assert.Equal("Operator", info.Operator);
        Assert.Equal("Category", info.Category);
        Assert.Equal("one,two", info.Keywords);
        Assert.Equal("Comment", info.Comments);
        Assert.Equal("https://example.test/", info.HyperlinkBase);
        Assert.Equal(new DateTime(2026, 6, 15, 10, 20, 30), info.Created);
        Assert.Equal(new DateTime(2026, 6, 16, 11, 21, 31), info.Revised);
        Assert.Equal(42, info.EditingMinutes);
        Assert.Equal(7, info.NumberOfPages);
        Assert.Equal(120, info.NumberOfWords);
        Assert.Equal(600, info.NumberOfCharacters);
        Assert.Equal(700, info.NumberOfCharactersWithSpaces);
        Assert.Equal(123, info.InternalVersion);

        string written = result.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfDocumentInfo writtenInfo = RtfDocument.Read(written).Document.Info;

        Assert.Contains(@"{\hlinkbase https://example.test/}", written, StringComparison.Ordinal);
        Assert.Contains(@"{\creatim\yr2026\mo6\dy15\hr10\min20\sec30}", written, StringComparison.Ordinal);
        Assert.Contains(@"\edmins42", written, StringComparison.Ordinal);
        Assert.Equal(info.Title, writtenInfo.Title);
        Assert.Equal(info.Created, writtenInfo.Created);
        Assert.Equal(info.Revised, writtenInfo.Revised);
        Assert.Equal(info.NumberOfCharactersWithSpaces, writtenInfo.NumberOfCharactersWithSpaces);
        Assert.Equal(info.InternalVersion, writtenInfo.InternalVersion);
    }

    [Fact]
    public void Read_Binds_Generator_Metadata_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\generator RichEdit 10.0.19041;}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal("RichEdit 10.0.19041", result.Document.Info.Generator);
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_User_Properties_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\userprops{\propname Client}\proptype30{\staticval Contoso}{\propname Approved}\proptype11{\staticval 1}{\propname Score}\proptype5{\staticval 98.5}{\propname External}{\linkval Sheet1!A1}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(4, result.Document.UserProperties.Count);
        Assert.Collection(result.Document.UserProperties,
            property => {
                Assert.Equal("Client", property.Name);
                Assert.Equal(RtfUserProperty.TextType, property.TypeCode);
                Assert.Equal("Contoso", property.StaticValue);
                Assert.Null(property.LinkedValue);
            },
            property => {
                Assert.Equal("Approved", property.Name);
                Assert.Equal(RtfUserProperty.BooleanType, property.TypeCode);
                Assert.Equal("1", property.StaticValue);
            },
            property => {
                Assert.Equal("Score", property.Name);
                Assert.Equal(RtfUserProperty.NumberType, property.TypeCode);
                Assert.Equal("98.5", property.StaticValue);
            },
            property => {
                Assert.Equal("External", property.Name);
                Assert.Null(property.TypeCode);
                Assert.Null(property.StaticValue);
                Assert.Equal("Sheet1!A1", property.LinkedValue);
            });
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Document_Variables_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\docvar {Client}{Contoso}}{\*\docvar {Region}{EMEA}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.DocumentVariables,
            variable => {
                Assert.Equal("Client", variable.Name);
                Assert.Equal("Contoso", variable.Value);
            },
            variable => {
                Assert.Equal("Region", variable.Name);
                Assert.Equal("EMEA", variable.Value);
            });
        Assert.Equal("Body", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Run_Revisions_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\revtbl{Alice;}{Bob;}}\pard Base {\revised\revauth0\revdttm123 Inserted\revised0} {\deleted\revauth1 Removed\deleted0}\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Collection(result.Document.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Base Inserted Removed", paragraph.ToPlainText());
        Assert.Contains(paragraph.Runs, run =>
            run.Text.Contains("Inserted", StringComparison.Ordinal) &&
            run.RevisionKind == RtfRevisionKind.Inserted &&
            run.RevisionAuthorIndex == 0 &&
            run.RevisionTimestampValue == 123);
        Assert.Contains(paragraph.Runs, run =>
            run.Text.Contains("Removed", StringComparison.Ordinal) &&
            run.RevisionKind == RtfRevisionKind.Deleted &&
            run.RevisionAuthorIndex == 1);
    }

    [Fact]
    public void Read_Binds_Revision_Save_Id_Table_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\*\rsidtbl\rsidroot7\rsid15\rsid1024\rsid65535}\pard\pararsid20 Base \charrsid30\insrsid40\delrsid50 Revised\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(7, result.Document.RevisionRootSaveId);
        Assert.Equal(new[] { 15, 1024, 65535 }, result.Document.RevisionSaveIds);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(20, paragraph.RevisionSaveId);
        Assert.Equal("Base Revised", paragraph.ToPlainText());
        RtfRun revised = paragraph.Runs.Single(run => run.Text == "Revised");
        Assert.Equal(30, revised.CharacterRevisionSaveId);
        Assert.Equal(40, revised.InsertionRevisionSaveId);
        Assert.Equal(50, revised.DeletionRevisionSaveId);
    }

    [Fact]
    public void Read_Binds_Annotation_Metadata_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Target{\annotation{\*\atnid c1}{\*\atnauthor Alice}{\*\atntime\yr2026\mo1\dy2\hr3\min4\sec5}\chatn\pard Review note\par}\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        RtfRun run = Assert.Single(paragraph.Runs);
        Assert.Equal("Target", run.Text);
        Assert.NotNull(run.Note);
        RtfNote note = run.Note!;
        Assert.Equal(RtfNoteKind.Annotation, note.Kind);
        Assert.Equal("c1", note.Id);
        Assert.Equal("Alice", note.Author);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), note.Created);
        Assert.Equal("Review note", note.ToPlainText());
    }
}
