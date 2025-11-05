using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private SdtBlock CoverPageBanded {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId();

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "002E223F", RsidRunAdditionDefault = "007522C7", ParagraphId = "65DA364D", TextId = "74CB0667" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "17D0774C" };

                V.Group group1 = new V.Group() { Id = "Group 193", Style = "position:absolute;margin-left:0;margin-top:0;width:540.55pt;height:718.4pt;z-index:-251657216;mso-width-percent:882;mso-height-percent:909;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:882;mso-height-percent:909", CoordinateSize = "68648,91235", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nBgAAAAAEAAQA8wAAAA4HAAAAAA==\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 194", Style = "position:absolute;width:68580;height:13716;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1027", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDHrpG1xAAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Na8JA\nEL0L/Q/LFHozm1ZpNbqKCEIREUzrwduQnWbTZmdDdhujv94VCr3N433OfNnbWnTU+sqxguckBUFc\nOF1xqeDzYzOcgPABWWPtmBRcyMNy8TCYY6bdmQ/U5aEUMYR9hgpMCE0mpS8MWfSJa4gj9+VaiyHC\ntpS6xXMMt7V8SdNXabHi2GCwobWh4if/tQq232+j3HSr7jra09G44+60WXulnh771QxEoD78i//c\n7zrOn47h/ky8QC5uAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAMeukbXEAAAA3AAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 195", Style = "position:absolute;top:40943;width:68580;height:50292;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", OptionalString = "_x0000_s1028", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBhwpCDxAAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Na8JA\nEL0L/Q/LFLwU3ShtqamriCLUIsXGXLyN2Wk2mJ0N2VXjv+8WCt7m8T5nOu9sLS7U+sqxgtEwAUFc\nOF1xqSDfrwdvIHxA1lg7JgU38jCfPfSmmGp35W+6ZKEUMYR9igpMCE0qpS8MWfRD1xBH7se1FkOE\nbSl1i9cYbms5TpJXabHi2GCwoaWh4pSdrYIsX+VHCs+Tz6/Dxu3yJ7Pbjjul+o/d4h1EoC7cxf/u\nDx3nT17g75l4gZz9AgAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAGHCkIPEAAAA3AAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

                V.TextBox textBox1 = new V.TextBox() { Inset = "36pt,57.6pt,36pt,36pt" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties2.Append(color1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Author" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId();
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(tag1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "002E223F", RsidRunAdditionDefault = "007522C7", ParagraphId = "17E41E5E", TextId = "36A9CA2E" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "120" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties1.Append(color2);

                paragraphProperties1.Append(spacingBetweenLines1);
                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties3.Append(color3);
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = "     ";

                run2.Append(runProperties3);
                run2.Append(text1);

                paragraph2.Append(paragraphProperties1);
                paragraph2.Append(run2);

                sdtContentBlock2.Append(paragraph2);

                sdtBlock2.Append(sdtProperties2);
                sdtBlock2.Append(sdtEndCharProperties2);
                sdtBlock2.Append(sdtContentBlock2);

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "002E223F", RsidRunAdditionDefault = "007522C7", ParagraphId = "256DC1E9", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "120" };
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties2.Append(color4);

                paragraphProperties2.Append(spacingBetweenLines2);
                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Caps caps1 = new Caps();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties4.Append(caps1);
                runProperties4.Append(color5);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Company" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId();
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\' ", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties4);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Caps caps2 = new Caps();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties5.Append(caps2);
                runProperties5.Append(color6);
                Text text2 = new Text();
                text2.Text = "[Company name]";

                run3.Append(runProperties5);
                run3.Append(text2);

                sdtContentRun1.Append(run3);

                sdtRun1.Append(sdtProperties3);
                sdtRun1.Append(sdtEndCharProperties3);
                sdtRun1.Append(sdtContentRun1);

                Run run4 = new Run();

                RunProperties runProperties6 = new RunProperties();
                Color color7 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties6.Append(color7);
                Text text3 = new Text();
                text3.Text = "  ";

                run4.Append(runProperties6);
                run4.Append(text3);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Color color8 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties7.Append(color8);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Address" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId();
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:CompanyAddress[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties7);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run5 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Color color9 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties8.Append(color9);
                Text text4 = new Text();
                text4.Text = "[Company address]";

                run5.Append(runProperties8);
                run5.Append(text4);

                sdtContentRun2.Append(run5);

                sdtRun2.Append(sdtProperties4);
                sdtRun2.Append(sdtEndCharProperties4);
                sdtRun2.Append(sdtContentRun2);

                paragraph3.Append(paragraphProperties2);
                paragraph3.Append(sdtRun1);
                paragraph3.Append(run4);
                paragraph3.Append(sdtRun2);

                textBoxContent1.Append(sdtBlock2);
                textBoxContent1.Append(paragraph3);

                textBox1.Append(textBoxContent1);

                rectangle2.Append(textBox1);

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 196", Style = "position:absolute;left:68;top:13716;width:68580;height:27227;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1029", FillColor = "white [3212]", Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCT/NOqwgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Li8Iw\nEL4L/ocwghdZ07WgazWKD8T1qC4s3oZmbIvNpNtErf/eCAve5uN7znTemFLcqHaFZQWf/QgEcWp1\nwZmCn+Pm4wuE88gaS8uk4EEO5rN2a4qJtnfe0+3gMxFC2CWoIPe+SqR0aU4GXd9WxIE729qgD7DO\npK7xHsJNKQdRNJQGCw4NOVa0yim9HK5GwXjp93Hv9xRX2z+zxuy6O8ajk1LdTrOYgPDU+Lf43/2t\nw/zxEF7PhAvk7AkAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQCT/NOqwgAAANwAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };

                V.TextBox textBox2 = new V.TextBox() { Inset = "36pt,7.2pt,36pt,7.2pt" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps3 = new Caps();
                Color color10 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize1 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "72" };

                runProperties9.Append(runFonts1);
                runProperties9.Append(caps3);
                runProperties9.Append(color10);
                runProperties9.Append(fontSize1);
                runProperties9.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Title" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId();
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties9);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "002E223F", RsidRunAdditionDefault = "007522C7", ParagraphId = "77197ED1", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification3 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps4 = new Caps();
                Color color11 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize2 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "72" };

                paragraphMarkRunProperties3.Append(runFonts2);
                paragraphMarkRunProperties3.Append(caps4);
                paragraphMarkRunProperties3.Append(color11);
                paragraphMarkRunProperties3.Append(fontSize2);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript2);

                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run6 = new Run();

                RunProperties runProperties10 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps5 = new Caps();
                Color color12 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize3 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "72" };

                runProperties10.Append(runFonts3);
                runProperties10.Append(caps5);
                runProperties10.Append(color12);
                runProperties10.Append(fontSize3);
                runProperties10.Append(fontSizeComplexScript3);
                Text text5 = new Text();
                text5.Text = "[Document title]";

                run6.Append(runProperties10);
                run6.Append(text5);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(run6);

                sdtContentBlock3.Append(paragraph4);

                sdtBlock3.Append(sdtProperties5);
                sdtBlock3.Append(sdtEndCharProperties5);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent2.Append(sdtBlock3);

                textBox2.Append(textBoxContent2);

                shape1.Append(textBox2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                group1.Append(rectangle1);
                group1.Append(rectangle2);
                group1.Append(shapetype1);
                group1.Append(shape1);
                group1.Append(textWrap1);

                picture1.Append(group1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(run1);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "002E223F", RsidRunAdditionDefault = "007522C7", ParagraphId = "19D15993", TextId = "49FFEE6A" };

                Run run7 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run7.Append(break1);

                paragraph5.Append(run7);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph5);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }
    }
}
