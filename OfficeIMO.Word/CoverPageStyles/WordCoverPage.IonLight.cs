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
    public partial class WordCoverPage {

        private static SdtBlock CoverPageIonLight {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 662514150 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();
                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "76EAEF43", TextId = "53BCF98E" };

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "3FD3FBEA", TextId = "6037CDFD" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "0870C5C2" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 131", Style = "position:absolute;margin-left:0;margin-top:0;width:369pt;height:529.2pt;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-width-percent:790;mso-height-percent:350;mso-left-percent:77;mso-top-percent:540;mso-wrap-distance-left:14.4pt;mso-wrap-distance-top:0;mso-wrap-distance-right:14.4pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:margin;mso-position-vertical-relative:page;mso-width-percent:790;mso-height-percent:350;mso-left-percent:77;mso-top-percent:540;mso-width-relative:margin;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\ncmV2LnhtbFBLBQYAAAAABAAEAPMAAADCBQAAAAA=\n" };

                V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "4F783CCE", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "560", Line = "216", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color1 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize1 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "72" };

                paragraphMarkRunProperties1.Append(color1);
                paragraphMarkRunProperties1.Append(fontSize1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties1.Append(spacingBetweenLines1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Color color2 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize2 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "72" };

                runProperties2.Append(color2);
                runProperties2.Append(fontSize2);
                runProperties2.Append(fontSizeComplexScript2);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId() { Val = 151731938 };
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(tag1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Color color3 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize3 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "72" };

                runProperties3.Append(color3);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "[Document title]";

                run2.Append(runProperties3);
                run2.Append(text1);

                sdtContentRun1.Append(run2);

                sdtRun1.Append(sdtProperties2);
                sdtRun1.Append(sdtEndCharProperties2);
                sdtRun1.Append(sdtContentRun1);

                paragraph3.Append(paragraphProperties1);
                paragraph3.Append(sdtRun1);

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Caps caps1 = new Caps();
                Color color4 = new Color() { Val = "1F4E79", ThemeColor = ThemeColorValues.Accent5, ThemeShade = "80" };
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties4.Append(caps1);
                runProperties4.Append(color4);
                runProperties4.Append(fontSize4);
                runProperties4.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Subtitle" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId() { Val = -2090151685 };
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties4);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "52C186FF", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "40", After = "40" };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Caps caps2 = new Caps();
                Color color5 = new Color() { Val = "1F4E79", ThemeColor = ThemeColorValues.Accent5, ThemeShade = "80" };
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties2.Append(caps2);
                paragraphMarkRunProperties2.Append(color5);
                paragraphMarkRunProperties2.Append(fontSize5);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript5);

                paragraphProperties2.Append(spacingBetweenLines2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Caps caps3 = new Caps();
                Color color6 = new Color() { Val = "1F4E79", ThemeColor = ThemeColorValues.Accent5, ThemeShade = "80" };
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(caps3);
                runProperties5.Append(color6);
                runProperties5.Append(fontSize6);
                runProperties5.Append(fontSizeComplexScript6);
                Text text2 = new Text();
                text2.Text = "[Document subtitle]";

                run3.Append(runProperties5);
                run3.Append(text2);

                paragraph4.Append(paragraphProperties2);
                paragraph4.Append(run3);

                sdtContentBlock2.Append(paragraph4);

                sdtBlock2.Append(sdtProperties3);
                sdtBlock2.Append(sdtEndCharProperties3);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties6 = new RunProperties();
                Caps caps4 = new Caps();
                Color color7 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize7 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

                runProperties6.Append(caps4);
                runProperties6.Append(color7);
                runProperties6.Append(fontSize7);
                runProperties6.Append(fontSizeComplexScript7);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Author" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId() { Val = -1536112409 };
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties6);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "6564CE82", TextId = "176E7E5F" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "80", After = "40" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Caps caps5 = new Caps();
                Color color8 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize8 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties3.Append(caps5);
                paragraphMarkRunProperties3.Append(color8);
                paragraphMarkRunProperties3.Append(fontSize8);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript8);

                paragraphProperties3.Append(spacingBetweenLines3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run4 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Caps caps6 = new Caps();
                Color color9 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize9 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

                runProperties7.Append(caps6);
                runProperties7.Append(color9);
                runProperties7.Append(fontSize9);
                runProperties7.Append(fontSizeComplexScript9);
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = "     ";

                run4.Append(runProperties7);
                run4.Append(text3);

                paragraph5.Append(paragraphProperties3);
                paragraph5.Append(run4);

                sdtContentBlock3.Append(paragraph5);

                sdtBlock3.Append(sdtProperties4);
                sdtBlock3.Append(sdtEndCharProperties4);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent1.Append(paragraph3);
                textBoxContent1.Append(sdtBlock2);
                textBoxContent1.Append(sdtBlock3);

                textBox1.Append(textBoxContent1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape1.Append(textBox1);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                Run run5 = new Run();

                RunProperties runProperties8 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties8.Append(noProof2);

                Picture picture2 = new Picture() { AnchorId = "26546919" };

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 132", Style = "position:absolute;margin-left:-8.8pt;margin-top:0;width:46.8pt;height:77.75pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:76;mso-height-percent:98;mso-top-percent:23;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:right;mso-position-horizontal-relative:margin;mso-position-vertical-relative:page;mso-width-percent:76;mso-height-percent:98;mso-top-percent:23;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1027", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n5AUAAAAA\n"));
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                V.TextBox textBox2 = new V.TextBox() { Inset = "3.6pt,,3.6pt" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                Color color10 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize10 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

                runProperties9.Append(color10);
                runProperties9.Append(fontSize10);
                runProperties9.Append(fontSizeComplexScript10);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Year" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId() { Val = -785116381 };
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };

                SdtContentDate sdtContentDate1 = new SdtContentDate();
                DateFormat dateFormat1 = new DateFormat() { Val = "yyyy" };
                LanguageId languageId1 = new LanguageId() { Val = "en-US" };
                SdtDateMappingType sdtDateMappingType1 = new SdtDateMappingType() { Val = DateFormatValues.DateTime };
                Calendar calendar1 = new Calendar() { Val = CalendarValues.Gregorian };

                sdtContentDate1.Append(dateFormat1);
                sdtContentDate1.Append(languageId1);
                sdtContentDate1.Append(sdtDateMappingType1);
                sdtContentDate1.Append(calendar1);

                sdtProperties5.Append(runProperties9);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentDate1);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00644C41", RsidRunAdditionDefault = "007158FF", ParagraphId = "0070B145", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color11 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize11 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties4.Append(color11);
                paragraphMarkRunProperties4.Append(fontSize11);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript11);

                paragraphProperties4.Append(justification1);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run6 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Color color12 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize12 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

                runProperties10.Append(color12);
                runProperties10.Append(fontSize12);
                runProperties10.Append(fontSizeComplexScript12);
                Text text4 = new Text();
                text4.Text = "[Year]";

                run6.Append(runProperties10);
                run6.Append(text4);

                paragraph6.Append(paragraphProperties4);
                paragraph6.Append(run6);

                sdtContentBlock4.Append(paragraph6);

                sdtBlock4.Append(sdtProperties5);
                sdtBlock4.Append(sdtEndCharProperties5);
                sdtBlock4.Append(sdtContentBlock4);

                textBoxContent2.Append(sdtBlock4);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle1.Append(lock1);
                rectangle1.Append(textBox2);
                rectangle1.Append(textWrap2);

                picture2.Append(rectangle1);

                run5.Append(runProperties8);
                run5.Append(picture2);

                Run run7 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run7.Append(break1);

                paragraph2.Append(run1);
                paragraph2.Append(run5);
                paragraph2.Append(run7);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;


            }
        }
    }
}
