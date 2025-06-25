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

        private static SdtBlock CoverPageRetrospect {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 670766965 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "6ED7CF65", TextId = "2B8B2456" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "34344CFD" };

                V.Group group1 = new V.Group() { Id = "Group 119", Style = "position:absolute;margin-left:0;margin-top:0;width:539.6pt;height:719.9pt;z-index:-251657216;mso-width-percent:882;mso-height-percent:909;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:882;mso-height-percent:909", CoordinateSize = "68580,92717", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAAA=\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 120", Style = "position:absolute;top:73152;width:68580;height:1431;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1027", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAbKl5RxgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nEIXvBf/DMkJvdaNCK6mriCCUIoVGPfQ2ZMdsNDsbstsY++s7h0JvM7w3732zXA++UT11sQ5sYDrJ\nQBGXwdZcGTgedk8LUDEhW2wCk4E7RVivRg9LzG248Sf1RaqUhHDM0YBLqc21jqUjj3ESWmLRzqHz\nmGTtKm07vEm4b/Qsy561x5qlwWFLW0fltfj2Bt4vL/PC9Zv+Z/5BJxdO+6/dNhrzOB42r6ASDenf\n/Hf9ZgV/JvjyjEygV78AAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAA\nAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAA\nCwAAAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAGypeUcYAAADcAAAA\nDwAAAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPoCAAAAAA==\n"));

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 121", Style = "position:absolute;top:74390;width:68580;height:18327;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", OptionalString = "_x0000_s1028", FillColor = "#ed7d31 [3205]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBkMjDSwQAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Li8Iw\nEL4v+B/CCF4WTfXgSjWKCoplWfB5H5qxLW0mpYla//1GELzNx/ec2aI1lbhT4wrLCoaDCARxanXB\nmYLzadOfgHAeWWNlmRQ8ycFi3vmaYaztgw90P/pMhBB2MSrIva9jKV2ak0E3sDVx4K62MegDbDKp\nG3yEcFPJURSNpcGCQ0OONa1zSsvjzSjY/a7S4qc68L7clttLkkySv2+nVK/bLqcgPLX+I367dzrM\nHw3h9Uy4QM7/AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAA\nAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAA\nAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAGQyMNLBAAAA3AAAAA8AAAAA\nAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD1AgAAAAA=\n"));

                V.TextBox textBox1 = new V.TextBox() { Inset = "36pt,14.4pt,36pt,36pt" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Author" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId() { Val = 884141857 };
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

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "4CF111ED", TextId = "02B7A6B4" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

                runProperties3.Append(color3);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
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

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "574986AA", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Caps caps1 = new Caps();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties2.Append(caps1);
                paragraphMarkRunProperties2.Append(color4);

                paragraphProperties2.Append(paragraphMarkRunProperties2);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Caps caps2 = new Caps();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties4.Append(caps2);
                runProperties4.Append(color5);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Company" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId() { Val = 922067218 };
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
                Caps caps3 = new Caps();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties5.Append(caps3);
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
                Caps caps4 = new Caps();
                Color color7 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties6.Append(caps4);
                runProperties6.Append(color7);
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = " | ";

                run4.Append(runProperties6);
                run4.Append(text3);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Caps caps5 = new Caps();
                Color color8 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties7.Append(caps5);
                runProperties7.Append(color8);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Address" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId() { Val = 2113163453 };
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
                Caps caps6 = new Caps();
                Color color9 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties8.Append(caps6);
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

                V.Shape shape1 = new V.Shape() { Id = "Text Box 122", Style = "position:absolute;width:68580;height:73152;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDm6hvlwgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Na8JA\nEL0X+h+WKXirm4YiJbqKiEKhXjSiHsfsmA1mZ0N2NWl/vSsUvM3jfc5k1tta3Kj1lWMFH8MEBHHh\ndMWlgl2+ev8C4QOyxtoxKfglD7Pp68sEM+063tBtG0oRQ9hnqMCE0GRS+sKQRT90DXHkzq61GCJs\nS6lb7GK4rWWaJCNpseLYYLChhaHisr1aBatDf+L872dnjsvlZ3c9FbzP10oN3vr5GESgPjzF/+5v\nHeenKTyeiRfI6R0AAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDm6hvlwgAAANwAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };

                V.TextBox textBox2 = new V.TextBox() { Inset = "36pt,36pt,36pt,36pt" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color10 = new Color() { Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };
                FontSize fontSize4 = new FontSize() { Val = "108" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "108" };

                runProperties9.Append(runFonts1);
                runProperties9.Append(color10);
                runProperties9.Append(fontSize4);
                runProperties9.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Title" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId() { Val = -1476986296 };
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

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "75883F8A", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();

                ParagraphBorders paragraphBorders1 = new ParagraphBorders();
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80", Size = (UInt32Value)6U, Space = (UInt32Value)4U };

                paragraphBorders1.Append(bottomBorder1);

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color11 = new Color() { Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };
                FontSize fontSize5 = new FontSize() { Val = "108" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "108" };

                paragraphMarkRunProperties3.Append(runFonts2);
                paragraphMarkRunProperties3.Append(color11);
                paragraphMarkRunProperties3.Append(fontSize5);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

                paragraphProperties3.Append(paragraphBorders1);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run6 = new Run();

                RunProperties runProperties10 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color12 = new Color() { Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };
                FontSize fontSize6 = new FontSize() { Val = "108" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "108" };

                runProperties10.Append(runFonts3);
                runProperties10.Append(color12);
                runProperties10.Append(fontSize6);
                runProperties10.Append(fontSizeComplexScript6);
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

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties6 = new SdtProperties();

                RunProperties runProperties11 = new RunProperties();
                Caps caps7 = new Caps();
                Color color13 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize7 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "36" };

                runProperties11.Append(caps7);
                runProperties11.Append(color13);
                runProperties11.Append(fontSize7);
                runProperties11.Append(fontSizeComplexScript7);
                SdtAlias sdtAlias5 = new SdtAlias() { Val = "Subtitle" };
                Tag tag5 = new Tag() { Val = "" };
                SdtId sdtId6 = new SdtId() { Val = 157346227 };
                ShowingPlaceholder showingPlaceholder5 = new ShowingPlaceholder();
                DataBinding dataBinding5 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText5 = new SdtContentText();

                sdtProperties6.Append(runProperties11);
                sdtProperties6.Append(sdtAlias5);
                sdtProperties6.Append(tag5);
                sdtProperties6.Append(sdtId6);
                sdtProperties6.Append(showingPlaceholder5);
                sdtProperties6.Append(dataBinding5);
                sdtProperties6.Append(sdtContentText5);
                SdtEndCharProperties sdtEndCharProperties6 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "06FA96BC", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240" };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Caps caps8 = new Caps();
                Color color14 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize8 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties4.Append(caps8);
                paragraphMarkRunProperties4.Append(color14);
                paragraphMarkRunProperties4.Append(fontSize8);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript8);

                paragraphProperties4.Append(spacingBetweenLines1);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run7 = new Run();

                RunProperties runProperties12 = new RunProperties();
                Caps caps9 = new Caps();
                Color color15 = new Color() { Val = "44546A", ThemeColor = ThemeColorValues.Text2 };
                FontSize fontSize9 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "36" };

                runProperties12.Append(caps9);
                runProperties12.Append(color15);
                runProperties12.Append(fontSize9);
                runProperties12.Append(fontSizeComplexScript9);
                Text text6 = new Text();
                text6.Text = "[Document subtitle]";

                run7.Append(runProperties12);
                run7.Append(text6);

                paragraph5.Append(paragraphProperties4);
                paragraph5.Append(run7);

                sdtContentBlock4.Append(paragraph5);

                sdtBlock4.Append(sdtProperties6);
                sdtBlock4.Append(sdtEndCharProperties6);
                sdtBlock4.Append(sdtContentBlock4);

                textBoxContent2.Append(sdtBlock3);
                textBoxContent2.Append(sdtBlock4);

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

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00C31810", RsidRunAdditionDefault = "00130B52", ParagraphId = "19AECCE7", TextId = "57EA3220" };

                Run run8 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run8.Append(break1);

                paragraph6.Append(run8);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph6);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }
    }
}
