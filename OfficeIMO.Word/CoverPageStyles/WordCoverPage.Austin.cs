using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private SdtBlock CoverPageAustin {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1961918155 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "38709928", TextId = "2C5B026B" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "00F125BE" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 465", Style = "position:absolute;margin-left:0;margin-top:0;width:220.3pt;height:21.15pt;z-index:251664384;visibility:visible;mso-wrap-style:square;mso-width-percent:360;mso-height-percent:0;mso-left-percent:455;mso-top-percent:660;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:360;mso-height-percent:0;mso-left-percent:455;mso-top-percent:660;mso-width-relative:page;mso-height-relative:margin;v-text-anchor:bottom", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\nbFBLBQYAAAAABAAEAPMAAACDBQAAAAA=\n" };

                V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "0A89CF34", TextId = "23CA8925" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                NoProof noProof2 = new NoProof();
                Color color1 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };

                paragraphMarkRunProperties1.Append(noProof2);
                paragraphMarkRunProperties1.Append(color1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color2 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };

                runProperties2.Append(noProof3);
                runProperties2.Append(color2);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Author" };
                SdtId sdtId2 = new SdtId() { Val = 15524260 };
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof4 = new NoProof();
                Color color3 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };

                runProperties3.Append(noProof4);
                runProperties3.Append(color3);
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = "     ";

                run2.Append(runProperties3);
                run2.Append(text1);

                sdtContentRun1.Append(run2);

                sdtRun1.Append(sdtProperties2);
                sdtRun1.Append(sdtEndCharProperties2);
                sdtRun1.Append(sdtContentRun1);

                paragraph2.Append(paragraphProperties1);
                paragraph2.Append(sdtRun1);

                textBoxContent1.Append(paragraph2);

                textBox1.Append(textBoxContent1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape1.Append(textBox1);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                Run run3 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof5 = new NoProof();

                runProperties4.Append(noProof5);

                Picture picture2 = new Picture() { AnchorId = "66BCDEFF" };

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 466", Style = "position:absolute;margin-left:0;margin-top:0;width:581.4pt;height:752.4pt;z-index:-251653120;visibility:visible;mso-wrap-style:square;mso-width-percent:950;mso-height-percent:950;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:950;mso-height-percent:950;mso-width-relative:page;mso-height-relative:page;v-text-anchor:middle", OptionalString = "_x0000_s1027", FillColor = "#d9e2f3 [660]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAAQA8wAAACQGAAAAAA==\n"));

                V.Fill fill1 = new V.Fill() { Type = V.FillTypeValues.Gradient, Color2 = "#8eaadb [1940]", Focus = "100%", Rotate = true };
                Ovml.FillExtendedProperties fillExtendedProperties1 = new Ovml.FillExtendedProperties() { Extension = V.ExtensionHandlingBehaviorValues.View, Type = Ovml.FillValues.GradientUnscaled };

                fill1.Append(fillExtendedProperties1);

                V.TextBox textBox2 = new V.TextBox() { Inset = "21.6pt,,21.6pt" };

                TextBoxContent textBoxContent2 = new TextBoxContent();
                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "6F27ED43", TextId = "77777777" };

                textBoxContent2.Append(paragraph3);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle1.Append(fill1);
                rectangle1.Append(textBox2);
                rectangle1.Append(textWrap2);

                picture2.Append(rectangle1);

                run3.Append(runProperties4);
                run3.Append(picture2);

                Run run4 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof6 = new NoProof();

                runProperties5.Append(noProof6);

                Picture picture3 = new Picture() { AnchorId = "57068562" };

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 467", Style = "position:absolute;margin-left:0;margin-top:0;width:226.45pt;height:237.6pt;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-width-percent:370;mso-height-percent:300;mso-left-percent:455;mso-top-percent:25;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:370;mso-height-percent:300;mso-left-percent:455;mso-top-percent:25;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1028", FillColor = "#44546a [3215]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nBADzAAAA6gUAAAAA\n"));

                V.TextBox textBox3 = new V.TextBox() { Inset = "14.4pt,14.4pt,14.4pt,28.8pt" };

                TextBoxContent textBoxContent3 = new TextBoxContent();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "6D3AEC1C", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties2.Append(color4);

                paragraphProperties2.Append(spacingBetweenLines1);
                paragraphProperties2.Append(justification1);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties6 = new RunProperties();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties6.Append(color5);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Abstract" };
                SdtId sdtId3 = new SdtId() { Val = 8276291 };
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\'", XPath = "/ns0:CoverPageProperties[1]/ns0:Abstract[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties6);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run5 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties7.Append(color6);
                Text text2 = new Text();
                text2.Text = "[Draw your reader in with an engaging abstract. It is typically a short summary of the document. When you’re ready to add your content, just click here and start typing.]";

                run5.Append(runProperties7);
                run5.Append(text2);

                sdtContentRun2.Append(run5);

                sdtRun2.Append(sdtProperties3);
                sdtRun2.Append(sdtEndCharProperties3);
                sdtRun2.Append(sdtContentRun2);

                paragraph4.Append(paragraphProperties2);
                paragraph4.Append(sdtRun2);

                textBoxContent3.Append(paragraph4);

                textBox3.Append(textBoxContent3);
                Wvml.TextWrap textWrap3 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle2.Append(textBox3);
                rectangle2.Append(textWrap3);

                picture3.Append(rectangle2);

                run4.Append(runProperties5);
                run4.Append(picture3);

                Run run6 = new Run();

                RunProperties runProperties8 = new RunProperties();
                NoProof noProof7 = new NoProof();

                runProperties8.Append(noProof7);

                Picture picture4 = new Picture() { AnchorId = "6C141814" };

                V.Rectangle rectangle3 = new V.Rectangle() { Id = "Rectangle 468", Style = "position:absolute;margin-left:0;margin-top:0;width:244.8pt;height:554.4pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:400;mso-height-percent:700;mso-left-percent:440;mso-top-percent:25;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:400;mso-height-percent:700;mso-left-percent:440;mso-top-percent:25;mso-width-relative:page;mso-height-relative:page;v-text-anchor:middle", OptionalString = "_x0000_s1031", FillColor = "white [3212]", StrokeColor = "#747070 [1614]", StrokeWeight = "1.25pt" };
                rectangle3.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nbWxQSwUGAAAAAAQABADzAAAA9gUAAAAA\n"));
                Wvml.TextWrap textWrap4 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle3.Append(textWrap4);

                picture4.Append(rectangle3);

                run6.Append(runProperties8);
                run6.Append(picture4);

                Run run7 = new Run();

                RunProperties runProperties9 = new RunProperties();
                NoProof noProof8 = new NoProof();

                runProperties9.Append(noProof8);

                Picture picture5 = new Picture() { AnchorId = "687851A6" };

                V.Rectangle rectangle4 = new V.Rectangle() { Id = "Rectangle 469", Style = "position:absolute;margin-left:0;margin-top:0;width:226.45pt;height:9.35pt;z-index:251662336;visibility:visible;mso-wrap-style:square;mso-width-percent:370;mso-height-percent:0;mso-left-percent:455;mso-top-percent:690;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:370;mso-height-percent:0;mso-left-percent:455;mso-top-percent:690;mso-width-relative:page;mso-height-relative:margin;v-text-anchor:middle", OptionalString = "_x0000_s1030", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle4.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));
                Wvml.TextWrap textWrap5 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle4.Append(textWrap5);

                picture5.Append(rectangle4);

                run7.Append(runProperties9);
                run7.Append(picture5);

                Run run8 = new Run();

                RunProperties runProperties10 = new RunProperties();
                NoProof noProof9 = new NoProof();

                runProperties10.Append(noProof9);

                Picture picture6 = new Picture() { AnchorId = "49AED513" };

                V.Shape shape2 = new V.Shape() { Id = "Text Box 470", Style = "position:absolute;margin-left:0;margin-top:0;width:220.3pt;height:194.9pt;z-index:251661312;visibility:visible;mso-wrap-style:square;mso-width-percent:360;mso-height-percent:280;mso-left-percent:455;mso-top-percent:350;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:360;mso-height-percent:280;mso-left-percent:455;mso-top-percent:350;mso-width-relative:page;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\nbWxQSwUGAAAAAAQABADzAAAAhAUAAAAA\n" };

                V.TextBox textBox4 = new V.TextBox() { Style = "mso-fit-shape-to-text:t" };

                TextBoxContent textBoxContent4 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties11 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof10 = new NoProof();
                Color color7 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize1 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "72" };

                runProperties11.Append(runFonts1);
                runProperties11.Append(noProof10);
                runProperties11.Append(color7);
                runProperties11.Append(fontSize1);
                runProperties11.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Title" };
                SdtId sdtId4 = new SdtId() { Val = -958338334 };
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties11);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "552BCBDC", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof11 = new NoProof();
                Color color8 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize2 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "144" };

                paragraphMarkRunProperties3.Append(runFonts2);
                paragraphMarkRunProperties3.Append(noProof11);
                paragraphMarkRunProperties3.Append(color8);
                paragraphMarkRunProperties3.Append(fontSize2);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript2);

                paragraphProperties3.Append(spacingBetweenLines2);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run9 = new Run();

                RunProperties runProperties12 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof12 = new NoProof();
                Color color9 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize3 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "72" };

                runProperties12.Append(runFonts3);
                runProperties12.Append(noProof12);
                runProperties12.Append(color9);
                runProperties12.Append(fontSize3);
                runProperties12.Append(fontSizeComplexScript3);
                Text text3 = new Text();
                text3.Text = "[Document title]";

                run9.Append(runProperties12);
                run9.Append(text3);

                paragraph5.Append(paragraphProperties3);
                paragraph5.Append(run9);

                sdtContentBlock2.Append(paragraph5);

                sdtBlock2.Append(sdtProperties4);
                sdtBlock2.Append(sdtEndCharProperties4);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties13 = new RunProperties();
                RunFonts runFonts4 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof13 = new NoProof();
                Color color10 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize4 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "32" };

                runProperties13.Append(runFonts4);
                runProperties13.Append(noProof13);
                runProperties13.Append(color10);
                runProperties13.Append(fontSize4);
                runProperties13.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Subtitle" };
                SdtId sdtId5 = new SdtId() { Val = 15524255 };
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties13);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "0C389CFA", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                RunFonts runFonts5 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof14 = new NoProof();
                Color color11 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize5 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "40" };

                paragraphMarkRunProperties4.Append(runFonts5);
                paragraphMarkRunProperties4.Append(noProof14);
                paragraphMarkRunProperties4.Append(color11);
                paragraphMarkRunProperties4.Append(fontSize5);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript5);

                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run10 = new Run();

                RunProperties runProperties14 = new RunProperties();
                RunFonts runFonts6 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof15 = new NoProof();
                Color color12 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize6 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "32" };

                runProperties14.Append(runFonts6);
                runProperties14.Append(noProof15);
                runProperties14.Append(color12);
                runProperties14.Append(fontSize6);
                runProperties14.Append(fontSizeComplexScript6);
                Text text4 = new Text();
                text4.Text = "[Document subtitle]";

                run10.Append(runProperties14);
                run10.Append(text4);

                paragraph6.Append(paragraphProperties4);
                paragraph6.Append(run10);

                sdtContentBlock3.Append(paragraph6);

                sdtBlock3.Append(sdtProperties5);
                sdtBlock3.Append(sdtEndCharProperties5);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent4.Append(sdtBlock2);
                textBoxContent4.Append(sdtBlock3);

                textBox4.Append(textBoxContent4);
                Wvml.TextWrap textWrap6 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape2.Append(textBox4);
                shape2.Append(textWrap6);

                picture6.Append(shape2);

                run8.Append(runProperties10);
                run8.Append(picture6);

                paragraph1.Append(run1);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run6);
                paragraph1.Append(run7);
                paragraph1.Append(run8);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "001A2370", RsidRunAdditionDefault = "00FE22BC", ParagraphId = "3C003407", TextId = "3FEB9DC3" };

                Run run11 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run11.Append(break1);

                paragraph7.Append(run11);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph7);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }
    }
}
