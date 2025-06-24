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

        private static SdtBlock CoverPageSemaphore {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 389148313 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();
                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "3930EC80", TextId = "29390599" };

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "4113DFD5", TextId = "772F0DE0" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "7294575A" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 111", Style = "position:absolute;margin-left:0;margin-top:0;width:288.25pt;height:287.5pt;z-index:251662336;visibility:visible;mso-wrap-style:square;mso-width-percent:734;mso-height-percent:363;mso-left-percent:150;mso-top-percent:91;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:734;mso-height-percent:363;mso-left-percent:150;mso-top-percent:91;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

                V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Caps caps1 = new Caps();
                Color color1 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "40" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "40" };

                runProperties2.Append(caps1);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Publish Date" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId() { Val = 400952559 };
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };

                SdtContentDate sdtContentDate1 = new SdtContentDate();
                DateFormat dateFormat1 = new DateFormat() { Val = "MMMM d, yyyy" };
                LanguageId languageId1 = new LanguageId() { Val = "en-US" };
                SdtDateMappingType sdtDateMappingType1 = new SdtDateMappingType() { Val = DateFormatValues.DateTime };
                Calendar calendar1 = new Calendar() { Val = CalendarValues.Gregorian };

                sdtContentDate1.Append(dateFormat1);
                sdtContentDate1.Append(languageId1);
                sdtContentDate1.Append(sdtDateMappingType1);
                sdtContentDate1.Append(calendar1);

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(tag1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentDate1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "0B5984C4", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Caps caps2 = new Caps();
                Color color2 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize2 = new FontSize() { Val = "40" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "40" };

                paragraphMarkRunProperties1.Append(caps2);
                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Caps caps3 = new Caps();
                Color color3 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize3 = new FontSize() { Val = "40" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "40" };

                runProperties3.Append(caps3);
                runProperties3.Append(color3);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "[Date]";

                run2.Append(runProperties3);
                run2.Append(text1);

                paragraph3.Append(paragraphProperties1);
                paragraph3.Append(run2);

                sdtContentBlock2.Append(paragraph3);

                sdtBlock2.Append(sdtProperties2);
                sdtBlock2.Append(sdtEndCharProperties2);
                sdtBlock2.Append(sdtContentBlock2);

                textBoxContent1.Append(sdtBlock2);

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
                NoProof noProof2 = new NoProof();

                runProperties4.Append(noProof2);

                Picture picture2 = new Picture() { AnchorId = "17730A7A" };

                V.Shape shape2 = new V.Shape() { Id = "Text Box 112", Style = "position:absolute;margin-left:0;margin-top:0;width:453pt;height:51.4pt;z-index:251661312;visibility:visible;mso-wrap-style:square;mso-width-percent:734;mso-height-percent:80;mso-left-percent:150;mso-top-percent:837;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:734;mso-height-percent:80;mso-left-percent:150;mso-top-percent:837;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\ndi54bWxQSwUGAAAAAAQABADzAAAAwAUAAAAA\n" };

                V.TextBox textBox2 = new V.TextBox() { Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties5 = new RunProperties();
                Caps caps4 = new Caps();
                Color color4 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(caps4);
                runProperties5.Append(color4);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Author" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId() { Val = 1901796142 };
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties3.Append(runProperties5);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "34596D10", TextId = "57F28548" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Caps caps5 = new Caps();
                Color color5 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties2.Append(caps5);
                paragraphMarkRunProperties2.Append(color5);
                paragraphMarkRunProperties2.Append(fontSize5);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript5);

                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run4 = new Run();

                RunProperties runProperties6 = new RunProperties();
                Caps caps6 = new Caps();
                Color color6 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties6.Append(caps6);
                runProperties6.Append(color6);
                runProperties6.Append(fontSize6);
                runProperties6.Append(fontSizeComplexScript6);
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = "     ";

                run4.Append(runProperties6);
                run4.Append(text2);

                paragraph4.Append(paragraphProperties2);
                paragraph4.Append(run4);

                sdtContentBlock3.Append(paragraph4);

                sdtBlock3.Append(sdtProperties3);
                sdtBlock3.Append(sdtEndCharProperties3);
                sdtBlock3.Append(sdtContentBlock3);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "65B9DE92", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification3 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Caps caps7 = new Caps();
                Color color7 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize7 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties3.Append(caps7);
                paragraphMarkRunProperties3.Append(color7);
                paragraphMarkRunProperties3.Append(fontSize7);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript7);

                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Caps caps8 = new Caps();
                Color color8 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize8 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "20" };

                runProperties7.Append(caps8);
                runProperties7.Append(color8);
                runProperties7.Append(fontSize8);
                runProperties7.Append(fontSizeComplexScript8);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Company" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId() { Val = -661235724 };
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\' ", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties4.Append(runProperties7);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run5 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Caps caps9 = new Caps();
                Color color9 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize9 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };

                runProperties8.Append(caps9);
                runProperties8.Append(color9);
                runProperties8.Append(fontSize9);
                runProperties8.Append(fontSizeComplexScript9);
                Text text3 = new Text();
                text3.Text = "[Company name]";

                run5.Append(runProperties8);
                run5.Append(text3);

                sdtContentRun1.Append(run5);

                sdtRun1.Append(sdtProperties4);
                sdtRun1.Append(sdtEndCharProperties4);
                sdtRun1.Append(sdtContentRun1);

                paragraph5.Append(paragraphProperties3);
                paragraph5.Append(sdtRun1);

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "269C6620", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                Justification justification4 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Caps caps10 = new Caps();
                Color color10 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize10 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties4.Append(caps10);
                paragraphMarkRunProperties4.Append(color10);
                paragraphMarkRunProperties4.Append(fontSize10);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript10);

                paragraphProperties4.Append(justification4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                Color color11 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize11 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };

                runProperties9.Append(color11);
                runProperties9.Append(fontSize11);
                runProperties9.Append(fontSizeComplexScript11);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Address" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId() { Val = 171227497 };
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:CompanyAddress[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties5.Append(runProperties9);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run6 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Color color12 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize12 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "20" };

                runProperties10.Append(color12);
                runProperties10.Append(fontSize12);
                runProperties10.Append(fontSizeComplexScript12);
                Text text4 = new Text();
                text4.Text = "[Company address]";

                run6.Append(runProperties10);
                run6.Append(text4);

                sdtContentRun2.Append(run6);

                sdtRun2.Append(sdtProperties5);
                sdtRun2.Append(sdtEndCharProperties5);
                sdtRun2.Append(sdtContentRun2);

                Run run7 = new Run();

                RunProperties runProperties11 = new RunProperties();
                Color color13 = new Color() { Val = "262626", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D9" };
                FontSize fontSize13 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "20" };

                runProperties11.Append(color13);
                runProperties11.Append(fontSize13);
                runProperties11.Append(fontSizeComplexScript13);
                Text text5 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text5.Text = " ";

                run7.Append(runProperties11);
                run7.Append(text5);

                paragraph6.Append(paragraphProperties4);
                paragraph6.Append(sdtRun2);
                paragraph6.Append(run7);

                textBoxContent2.Append(sdtBlock3);
                textBoxContent2.Append(paragraph5);
                textBoxContent2.Append(paragraph6);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape2.Append(textBox2);
                shape2.Append(textWrap2);

                picture2.Append(shape2);

                run3.Append(runProperties4);
                run3.Append(picture2);

                Run run8 = new Run();

                RunProperties runProperties12 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties12.Append(noProof3);

                Picture picture3 = new Picture() { AnchorId = "4F594EA6" };

                V.Shape shape3 = new V.Shape() { Id = "Text Box 113", Style = "position:absolute;margin-left:0;margin-top:0;width:453pt;height:41.4pt;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-width-percent:734;mso-height-percent:363;mso-left-percent:150;mso-top-percent:455;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:734;mso-height-percent:363;mso-left-percent:150;mso-top-percent:455;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1028", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

                V.TextBox textBox3 = new V.TextBox() { Inset = "0,0,0,0" };

                TextBoxContent textBoxContent3 = new TextBoxContent();

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "6E42AEE1", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                Justification justification5 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                Caps caps11 = new Caps();
                Color color14 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize14 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "52" };

                paragraphMarkRunProperties5.Append(caps11);
                paragraphMarkRunProperties5.Append(color14);
                paragraphMarkRunProperties5.Append(fontSize14);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript14);

                paragraphProperties5.Append(justification5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                SdtRun sdtRun3 = new SdtRun();

                SdtProperties sdtProperties6 = new SdtProperties();

                RunProperties runProperties13 = new RunProperties();
                Caps caps12 = new Caps();
                Color color15 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize15 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "52" };

                runProperties13.Append(caps12);
                runProperties13.Append(color15);
                runProperties13.Append(fontSize15);
                runProperties13.Append(fontSizeComplexScript15);
                SdtAlias sdtAlias5 = new SdtAlias() { Val = "Title" };
                Tag tag5 = new Tag() { Val = "" };
                SdtId sdtId6 = new SdtId() { Val = -1315561441 };
                ShowingPlaceholder showingPlaceholder5 = new ShowingPlaceholder();
                DataBinding dataBinding5 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText() { MultiLine = true };

                sdtProperties6.Append(runProperties13);
                sdtProperties6.Append(sdtAlias5);
                sdtProperties6.Append(tag5);
                sdtProperties6.Append(sdtId6);
                sdtProperties6.Append(showingPlaceholder5);
                sdtProperties6.Append(dataBinding5);
                sdtProperties6.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties6 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun3 = new SdtContentRun();

                Run run9 = new Run();

                RunProperties runProperties14 = new RunProperties();
                Caps caps13 = new Caps();
                Color color16 = new Color() { Val = "323E4F", ThemeColor = ThemeColorValues.Text2, ThemeShade = "BF" };
                FontSize fontSize16 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "52" };

                runProperties14.Append(caps13);
                runProperties14.Append(color16);
                runProperties14.Append(fontSize16);
                runProperties14.Append(fontSizeComplexScript16);
                Text text6 = new Text();
                text6.Text = "[Document title]";

                run9.Append(runProperties14);
                run9.Append(text6);

                sdtContentRun3.Append(run9);

                sdtRun3.Append(sdtProperties6);
                sdtRun3.Append(sdtEndCharProperties6);
                sdtRun3.Append(sdtContentRun3);

                paragraph7.Append(paragraphProperties5);
                paragraph7.Append(sdtRun3);

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties7 = new SdtProperties();

                RunProperties runProperties15 = new RunProperties();
                SmallCaps smallCaps1 = new SmallCaps();
                Color color17 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize17 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "36" };

                runProperties15.Append(smallCaps1);
                runProperties15.Append(color17);
                runProperties15.Append(fontSize17);
                runProperties15.Append(fontSizeComplexScript17);
                SdtAlias sdtAlias6 = new SdtAlias() { Val = "Subtitle" };
                Tag tag6 = new Tag() { Val = "" };
                SdtId sdtId7 = new SdtId() { Val = 1615247542 };
                ShowingPlaceholder showingPlaceholder6 = new ShowingPlaceholder();
                DataBinding dataBinding6 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText5 = new SdtContentText();

                sdtProperties7.Append(runProperties15);
                sdtProperties7.Append(sdtAlias6);
                sdtProperties7.Append(tag6);
                sdtProperties7.Append(sdtId7);
                sdtProperties7.Append(showingPlaceholder6);
                sdtProperties7.Append(dataBinding6);
                sdtProperties7.Append(sdtContentText5);
                SdtEndCharProperties sdtEndCharProperties7 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00760994", RsidRunAdditionDefault = "00873FAF", ParagraphId = "0F3BD631", TextId = "77777777" };

                ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                Justification justification6 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                SmallCaps smallCaps2 = new SmallCaps();
                Color color18 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize18 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties6.Append(smallCaps2);
                paragraphMarkRunProperties6.Append(color18);
                paragraphMarkRunProperties6.Append(fontSize18);
                paragraphMarkRunProperties6.Append(fontSizeComplexScript18);

                paragraphProperties6.Append(justification6);
                paragraphProperties6.Append(paragraphMarkRunProperties6);

                Run run10 = new Run();

                RunProperties runProperties16 = new RunProperties();
                SmallCaps smallCaps3 = new SmallCaps();
                Color color19 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize19 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "36" };

                runProperties16.Append(smallCaps3);
                runProperties16.Append(color19);
                runProperties16.Append(fontSize19);
                runProperties16.Append(fontSizeComplexScript19);
                Text text7 = new Text();
                text7.Text = "[Document subtitle]";

                run10.Append(runProperties16);
                run10.Append(text7);

                paragraph8.Append(paragraphProperties6);
                paragraph8.Append(run10);

                sdtContentBlock4.Append(paragraph8);

                sdtBlock4.Append(sdtProperties7);
                sdtBlock4.Append(sdtEndCharProperties7);
                sdtBlock4.Append(sdtContentBlock4);

                textBoxContent3.Append(paragraph7);
                textBoxContent3.Append(sdtBlock4);

                textBox3.Append(textBoxContent3);
                Wvml.TextWrap textWrap3 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape3.Append(textBox3);
                shape3.Append(textWrap3);

                picture3.Append(shape3);

                run8.Append(runProperties12);
                run8.Append(picture3);

                Run run11 = new Run();

                RunProperties runProperties17 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties17.Append(noProof4);

                Picture picture4 = new Picture() { AnchorId = "618B2289" };

                V.Group group1 = new V.Group() { Id = "Group 114", Style = "position:absolute;margin-left:0;margin-top:0;width:18pt;height:10in;z-index:251659264;mso-width-percent:29;mso-height-percent:909;mso-left-percent:45;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:29;mso-height-percent:909;mso-left-percent:45", CoordinateSize = "2286,91440", OptionalString = "_x0000_s1029" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQDEPLGdIwMAAMUKAAAOAAAAZHJzL2Uyb0RvYy54bWzsVltP2zAUfp+0/2D5fSSpaCkRKarKQJMq\nQMDEs3GcJprj49luU/brd+xcgFLxwKZJk3gJvpzr1+98+OR0W0uyEcZWoDKaHMSUCMUhr9Qqo9/v\nzr9MKbGOqZxJUCKjj8LS09nnTyeNTsUISpC5MASDKJs2OqOlczqNIstLUTN7AFoovCzA1Mzh1qyi\n3LAGo9cyGsXxJGrA5NoAF9bi6Vl7SWchflEI7q6KwgpHZEaxNhe+Jnwf/DeanbB0ZZguK96Vwd5R\nRc0qhUmHUGfMMbI21atQdcUNWCjcAYc6gqKouAg9YDdJvNPNhYG1Dr2s0malB5gQ2h2c3h2WX24u\njL7V1waRaPQKsQg738u2MLX/i1WSbYDscYBMbB3heDgaTScxAsvx6jg5PIxxEzDlJQL/yo2XX992\njPq00YtiGo30sE8I2D9D4LZkWgRgbYoIXBtS5cjeZEyJYjXS9AaJw9RKCuIPAzTBcgDKphYxew9K\n06PpKB4HlIZmWaqNdRcCauIXGTWYP/CJbZbWYX407U18Uguyys8rKcPGD4tYSEM2DGnOOBfKjXzV\n6PXCUipvr8B7ttf+BKHu2wkr9yiFt5PqRhSIjP+ZQzFhKncTJe1VyXLR5h8jB/r2Bo9QSwjoIxeY\nf4jdBdjXRNI10dl7VxGGenCO3yqsbXHwCJlBucG5rhSYfQGkGzK39j1ILTQepQfIH5E3BlpJsZqf\nV/jTLZl118yghuBQoC66K/wUEpqMQreipATza9+5t0di4y0lDWpSRu3PNTOCEvlNIeXDgKGIhc3h\n+GiEOczzm4fnN2pdLwD5kKACax6W3t7JflkYqO9RPuc+K14xxTF3Rrkz/WbhWq1EAeZiPg9mKFya\nuaW61dwH96h6at5t75nRHX8d6sMl9GPG0h0at7beU8F87aCoAsefcO3wxpH3qvRPZn+yb/YnO7Pv\nS7Z6CfyHJQoWJWqEmFuN0+qh8Hzz1aKkeKFoS39TJ6bHyRg103sitfdIY6evLZN7Qe6V4K+JRc/2\nD7H4EIv/WyzCswHfSuH/Tfeu84+x5/swpU+vz9lvAAAA//8DAFBLAwQUAAYACAAAACEAvdF3w9oA\nAAAFAQAADwAAAGRycy9kb3ducmV2LnhtbEyPzU7DMBCE70h9B2srcaN2f1RBGqeqkOgNASkHenPi\nJYmw11HstuHtWbjQy0qjGc1+k29H78QZh9gF0jCfKRBIdbAdNRreD0939yBiMmSNC4QavjHCtpjc\n5Caz4UJveC5TI7iEYmY0tCn1mZSxbtGbOAs9EnufYfAmsRwaaQdz4XLv5EKptfSmI/7Qmh4fW6y/\nypPXQPJg97588R/L9FAujq+Ve95XWt9Ox90GRMIx/YfhF5/RoWCmKpzIRuE08JD0d9lbrllVnFmt\nlAJZ5PKavvgBAAD//wMAUEsBAi0AFAAGAAgAAAAhALaDOJL+AAAA4QEAABMAAAAAAAAAAAAAAAAA\nAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAOP0h/9YAAACUAQAACwAAAAAA\nAAAAAAAAAAAvAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAxDyxnSMDAADFCgAADgAAAAAA\nAAAAAAAAAAAuAgAAZHJzL2Uyb0RvYy54bWxQSwECLQAUAAYACAAAACEAvdF3w9oAAAAFAQAADwAA\nAAAAAAAAAAAAAAB9BQAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAAEAAQA8wAAAIQGAAAAAA==\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 115", Style = "position:absolute;width:2286;height:87820;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1031", FillColor = "#ed7d31 [3205]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCwN/CawAAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE/NisIw\nEL4v+A5hBG9r2oKi1SgqK8jiZasPMDZjW20mpcnW+vZmQdjbfHy/s1z3phYdta6yrCAeRyCIc6sr\nLhScT/vPGQjnkTXWlknBkxysV4OPJabaPviHuswXIoSwS1FB6X2TSunykgy6sW2IA3e1rUEfYFtI\n3eIjhJtaJlE0lQYrDg0lNrQrKb9nv0bBl7GT423emX1SXayczthvv1mp0bDfLEB46v2/+O0+6DA/\nnsDfM+ECuXoBAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAAAAAA\nAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAAAAAA\nAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAsDfwmsAAAADcAAAADwAAAAAA\nAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPQCAAAAAA==\n"));

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 116", Style = "position:absolute;top:89154;width:2286;height:2286;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1030", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQA146kDwgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Ni8Iw\nEL0v+B/CCN7WVAVXqlFEEBaRBasevA3N2FSbSWmyte6vNwsLe5vH+5zFqrOVaKnxpWMFo2ECgjh3\nuuRCwem4fZ+B8AFZY+WYFDzJw2rZe1tgqt2DD9RmoRAxhH2KCkwIdSqlzw1Z9ENXE0fu6hqLIcKm\nkLrBRwy3lRwnyVRaLDk2GKxpYyi/Z99Wwe72MclMu25/Jl90Nu68v2w3XqlBv1vPQQTqwr/4z/2p\n4/zRFH6fiRfI5QsAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQA146kDwgAAANwAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n"));
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                rectangle2.Append(lock1);
                Wvml.TextWrap textWrap4 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                group1.Append(rectangle1);
                group1.Append(rectangle2);
                group1.Append(textWrap4);

                picture4.Append(group1);

                run11.Append(runProperties17);
                run11.Append(picture4);

                Run run12 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run12.Append(break1);

                paragraph2.Append(run1);
                paragraph2.Append(run3);
                paragraph2.Append(run8);
                paragraph2.Append(run11);
                paragraph2.Append(run12);

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
