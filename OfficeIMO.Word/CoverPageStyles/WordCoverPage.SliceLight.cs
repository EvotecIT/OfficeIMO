using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageSliceLight {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();

                RunProperties runProperties1 = new RunProperties();
                FontSize fontSize1 = new FontSize() { Val = "2" };

                runProperties1.Append(fontSize1);
                SdtId sdtId1 = new SdtId();

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(runProperties1);
                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties2 = new RunProperties();
                FontSize fontSize2 = new FontSize() { Val = "22" };

                runProperties2.Append(fontSize2);

                sdtEndCharProperties1.Append(runProperties2);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "0DBAE5AB", TextId = "6B273C13" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                FontSize fontSize3 = new FontSize() { Val = "2" };

                paragraphMarkRunProperties1.Append(fontSize3);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                paragraph1.Append(paragraphProperties1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "56098EA3", TextId = "77777777" };

                Run run1 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties3.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "2422212B" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 62", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:1in;z-index:251661312;visibility:visible;mso-wrap-style:square;mso-width-percent:765;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:top;mso-position-vertical-relative:margin;mso-width-percent:765;mso-width-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

                V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps1 = new Caps();
                Color color1 = new Color() { Val = "8496B0", ThemeColor = ThemeColorValues.Text2, ThemeTint = "99" };
                FontSize fontSize4 = new FontSize() { Val = "64" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "64" };

                runProperties4.Append(runFonts1);
                runProperties4.Append(caps1);
                runProperties4.Append(color1);
                runProperties4.Append(fontSize4);
                runProperties4.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId();
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties4);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(tag1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);

                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                RunProperties runProperties5 = new RunProperties();
                FontSize fontSize5 = new FontSize() { Val = "68" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "68" };

                runProperties5.Append(fontSize5);
                runProperties5.Append(fontSizeComplexScript2);

                sdtEndCharProperties2.Append(runProperties5);

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "225AA849", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps2 = new Caps();
                Color color2 = new Color() { Val = "8496B0", ThemeColor = ThemeColorValues.Text2, ThemeTint = "99" };
                FontSize fontSize6 = new FontSize() { Val = "68" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "68" };

                paragraphMarkRunProperties2.Append(runFonts2);
                paragraphMarkRunProperties2.Append(caps2);
                paragraphMarkRunProperties2.Append(color2);
                paragraphMarkRunProperties2.Append(fontSize6);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run2 = new Run();

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps3 = new Caps();
                Color color3 = new Color() { Val = "8496B0", ThemeColor = ThemeColorValues.Text2, ThemeTint = "99" };
                FontSize fontSize7 = new FontSize() { Val = "68" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "68" };

                runProperties6.Append(runFonts3);
                runProperties6.Append(caps3);
                runProperties6.Append(color3);
                runProperties6.Append(fontSize7);
                runProperties6.Append(fontSizeComplexScript4);
                Text text1 = new Text();
                text1.Text = "[Document title]";

                run2.Append(runProperties6);
                run2.Append(text1);

                paragraph3.Append(paragraphProperties2);
                paragraph3.Append(run2);

                sdtContentBlock2.Append(paragraph3);

                sdtBlock2.Append(sdtProperties2);
                sdtBlock2.Append(sdtEndCharProperties2);
                sdtBlock2.Append(sdtContentBlock2);

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "0AF6BC42", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "120" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize8 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties3.Append(color4);
                paragraphMarkRunProperties3.Append(fontSize8);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

                paragraphProperties3.Append(spacingBetweenLines1);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Color color5 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize9 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "36" };

                runProperties7.Append(color5);
                runProperties7.Append(fontSize9);
                runProperties7.Append(fontSizeComplexScript6);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Subtitle" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId();
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties7);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run3 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Color color6 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize10 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "36" };

                runProperties8.Append(color6);
                runProperties8.Append(fontSize10);
                runProperties8.Append(fontSizeComplexScript7);
                Text text2 = new Text();
                text2.Text = "[Document subtitle]";

                run3.Append(runProperties8);
                run3.Append(text2);

                sdtContentRun1.Append(run3);

                sdtRun1.Append(sdtProperties3);
                sdtRun1.Append(sdtEndCharProperties3);
                sdtRun1.Append(sdtContentRun1);

                Run run4 = new Run();

                RunProperties runProperties9 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties9.Append(noProof2);
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = " ";

                run4.Append(runProperties9);
                run4.Append(text3);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(sdtRun1);
                paragraph4.Append(run4);
                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "36A0EB3F", TextId = "77777777" };

                textBoxContent1.Append(sdtBlock2);
                textBoxContent1.Append(paragraph4);
                textBoxContent1.Append(paragraph5);

                textBox1.Append(textBoxContent1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape1.Append(textBox1);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties3);
                run1.Append(picture1);

                Run run5 = new Run();

                RunProperties runProperties10 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color7 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize11 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "36" };

                runProperties10.Append(noProof3);
                runProperties10.Append(color7);
                runProperties10.Append(fontSize11);
                runProperties10.Append(fontSizeComplexScript8);

                Picture picture2 = new Picture() { AnchorId = "0A732961" };

                V.Group group1 = new V.Group() { Id = "Group 2", Style = "position:absolute;margin-left:0;margin-top:0;width:432.65pt;height:448.55pt;z-index:-251656192;mso-width-percent:706;mso-height-percent:566;mso-left-percent:220;mso-top-percent:300;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:706;mso-height-percent:566;mso-left-percent:220;mso-top-percent:300", CoordinateSize = "43291,44910", OptionalString = "_x0000_s1033" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                V.Shape shape2 = new V.Shape() { Id = "Freeform 64", Style = "position:absolute;left:15017;width:28274;height:28352;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "1781,1786", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, EdgePath = "m4,1786l,1782,1776,r5,5l4,1786xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDC/9XQwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx\nFITvQv9DeIXeNFspYlej2MK23mq3xfNj89wNbl62SVzXf28KgsdhZr5hluvBtqInH4xjBc+TDARx\n5bThWsHvTzGegwgRWWPrmBRcKMB69TBaYq7dmb+pL2MtEoRDjgqaGLtcylA1ZDFMXEecvIPzFmOS\nvpba4znBbSunWTaTFg2nhQY7em+oOpYnq6B/88NXdPttUZjdq+z1h/n73Cv19DhsFiAiDfEevrW3\nWsHsBf6/pB8gV1cAAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAwv/V0MMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };
                V.Path path2 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "6350,2835275;0,2828925;2819400,0;2827338,7938;6350,2835275", ConnectAngles = "0,0,0,0,0" };

                shape2.Append(path2);

                V.Shape shape3 = new V.Shape() { Id = "Freeform 65", Style = "position:absolute;left:7826;top:2270;width:35465;height:35464;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2234,2234", OptionalString = "_x0000_s1028", Filled = false, Stroked = false, EdgePath = "m5,2234l,2229,2229,r5,5l5,2234xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBNmpFYxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/RasJA\nFETfhf7Dcgt9001aDE10lRJa9EEsTfsBt9lrEpq9m2ZXjX69Kwg+DjNzhpkvB9OKA/WusawgnkQg\niEurG64U/Hx/jF9BOI+ssbVMCk7kYLl4GM0x0/bIX3QofCUChF2GCmrvu0xKV9Zk0E1sRxy8ne0N\n+iD7SuoejwFuWvkcRYk02HBYqLGjvKbyr9gbBcN5v9p8vsfdJmnTF/8r//N0i0o9PQ5vMxCeBn8P\n39prrSCZwvVL+AFycQEAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQBNmpFYxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n" };
                V.Path path3 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "7938,3546475;0,3538538;3538538,0;3546475,7938;7938,3546475", ConnectAngles = "0,0,0,0,0" };

                shape3.Append(path3);

                V.Shape shape4 = new V.Shape() { Id = "Freeform 66", Style = "position:absolute;left:8413;top:1095;width:34878;height:34877;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2197,2197", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, EdgePath = "m9,2197l,2193,2188,r9,10l9,2197xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAEeHK3xAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9La8Mw\nEITvhf4HsYVeSiKnB9c4kUMpuO01L0JuG2v9INbKtVTb/fdRINDjMDPfMKv1ZFoxUO8aywoW8wgE\ncWF1w5WC/S6fJSCcR9bYWiYFf+RgnT0+rDDVduQNDVtfiQBhl6KC2vsuldIVNRl0c9sRB6+0vUEf\nZF9J3eMY4KaVr1EUS4MNh4UaO/qoqbhsf42CxJ3Gtx3+fA5elovm5XzIj1+5Us9P0/sShKfJ/4fv\n7W+tII7h9iX8AJldAQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAAR4crfEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n" };
                V.Path path4 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "14288,3487738;0,3481388;3473450,0;3487738,15875;14288,3487738", ConnectAngles = "0,0,0,0,0" };

                shape4.Append(path4);

                V.Shape shape5 = new V.Shape() { Id = "Freeform 67", Style = "position:absolute;left:12160;top:4984;width:31131;height:31211;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "1961,1966", OptionalString = "_x0000_s1030", Filled = false, Stroked = false, EdgePath = "m9,1966l,1957,1952,r9,9l9,1966xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDdMg3hwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/BagIx\nEIbvBd8hjOCtZhXZ1tUoolSk0IO20Ou4mW6WbiZLkrrr2zeC4HH45//mm+W6t424kA+1YwWTcQaC\nuHS65krB1+fb8yuIEJE1No5JwZUCrFeDpyUW2nV8pMspViJBOBSowMTYFlKG0pDFMHYtccp+nLcY\n0+grqT12CW4bOc2yXFqsOV0w2NLWUPl7+rNJ43u628+MPCerPPs47uf+vZsrNRr2mwWISH18LN/b\nB60gf4HbLwkAcvUPAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEA3TIN4cMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };
                V.Path path5 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "14288,3121025;0,3106738;3098800,0;3113088,14288;14288,3121025", ConnectAngles = "0,0,0,0,0" };

                shape5.Append(path5);

                V.Shape shape6 = new V.Shape() { Id = "Freeform 68", Style = "position:absolute;top:1539;width:43291;height:43371;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2727,2732", OptionalString = "_x0000_s1031", Filled = false, Stroked = false, EdgePath = "m,2732r,-4l2722,r5,5l,2732xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBR/UPnuwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9LCsIw\nEN0L3iGM4E5TXZRSjaUIgi79HGBopm2wmZQmavX0ZiG4fLz/thhtJ540eONYwWqZgCCunDbcKLhd\nD4sMhA/IGjvHpOBNHorddLLFXLsXn+l5CY2IIexzVNCG0OdS+qoli37peuLI1W6wGCIcGqkHfMVw\n28l1kqTSouHY0GJP+5aq++VhFSRmferOaW20rLP7zZyyY/mplJrPxnIDItAY/uKf+6gVpHFs/BJ/\ngNx9AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAAAABb\nQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAAAAAA\nAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAFH9Q+e7AAAA2wAAAA8AAAAAAAAAAAAA\nAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADvAgAAAAA=\n" };
                V.Path path6 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "0,4337050;0,4330700;4321175,0;4329113,7938;0,4337050", ConnectAngles = "0,0,0,0,0" };

                shape6.Append(path6);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                group1.Append(lock1);
                group1.Append(shape2);
                group1.Append(shape3);
                group1.Append(shape4);
                group1.Append(shape5);
                group1.Append(shape6);
                group1.Append(textWrap2);

                picture2.Append(group1);

                run5.Append(runProperties10);
                run5.Append(picture2);

                Run run6 = new Run();

                RunProperties runProperties11 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties11.Append(noProof4);

                Picture picture3 = new Picture() { AnchorId = "7FB1D0C7" };

                V.Shape shape7 = new V.Shape() { Id = "Text Box 69", Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:29.5pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:765;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:bottom;mso-position-vertical-relative:margin;mso-width-percent:765;mso-height-percent:0;mso-width-relative:page;mso-height-relative:margin;v-text-anchor:bottom", OptionalString = "_x0000_s1032", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAYo6GiYQIAADQFAAAOAAAAZHJzL2Uyb0RvYy54bWysVEtv2zAMvg/YfxB0X+z0ka1BnSJLkWFA\n0RZrh54VWUqMyaJGKbGzXz9KtpOg26XDLjItfnx/1PVNWxu2U+grsAUfj3LOlJVQVnZd8O/Pyw+f\nOPNB2FIYsKrge+X5zez9u+vGTdUZbMCUChk5sX7auIJvQnDTLPNyo2rhR+CUJaUGrEWgX1xnJYqG\nvNcmO8vzSdYAlg5BKu/p9rZT8lnyr7WS4UFrrwIzBafcQjoxnat4ZrNrMV2jcJtK9mmIf8iiFpWl\noAdXtyIItsXqD1d1JRE86DCSUGegdSVVqoGqGeevqnnaCKdSLdQc7w5t8v/PrbzfPblHZKH9DC0N\nMDakcX7q6TLW02qs45cyZaSnFu4PbVNtYJIuL68uzic5qSTpzj9eXOUX0U12tHbowxcFNYtCwZHG\nkroldnc+dNABEoNZWFbGpNEYy5qCT84v82Rw0JBzYyNWpSH3bo6ZJynsjYoYY78pzaoyFRAvEr3U\nwiDbCSKGkFLZkGpPfgkdUZqSeIthjz9m9Rbjro4hMthwMK4rC5iqf5V2+WNIWXd46vlJ3VEM7aql\nwk8Gu4JyT/NG6FbBO7msaCh3wodHgcR9miPtc3igQxug5kMvcbYB/PW3+4gnSpKWs4Z2qeD+51ag\n4sx8tUTWuHiDgIOwGgS7rRdAUxjTS+FkEskAgxlEjVC/0JrPYxRSCSspVsFXg7gI3UbTMyHVfJ5A\ntF5OhDv75GR0HYcSKfbcvgh0PQ8DMfgehi0T01d07LCJL26+DUTKxNXY166Lfb9pNRPb+2ck7v7p\nf0IdH7vZbwAAAP//AwBQSwMEFAAGAAgAAAAhADHDoo3aAAAABAEAAA8AAABkcnMvZG93bnJldi54\nbWxMj91qwkAQhe8LvsMyhd7VTf9EYzYiUqGlFK31ASbZMQlmZ0N21fTtO+1NezNwOMM538kWg2vV\nmfrQeDZwN05AEZfeNlwZ2H+ub6egQkS22HomA18UYJGPrjJMrb/wB513sVISwiFFA3WMXap1KGty\nGMa+Ixbv4HuHUWRfadvjRcJdq++TZKIdNiwNNXa0qqk87k5OSsLxEPFx/f6mV8VLwc/b1+mmMubm\neljOQUUa4t8z/OALOuTCVPgT26BaAzIk/l7xZg8TkYWBp1kCOs/0f/j8GwAA//8DAFBLAQItABQA\nBgAIAAAAIQC2gziS/gAAAOEBAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1s\nUEsBAi0AFAAGAAgAAAAhADj9If/WAAAAlAEAAAsAAAAAAAAAAAAAAAAALwEAAF9yZWxzLy5yZWxz\nUEsBAi0AFAAGAAgAAAAhABijoaJhAgAANAUAAA4AAAAAAAAAAAAAAAAALgIAAGRycy9lMm9Eb2Mu\neG1sUEsBAi0AFAAGAAgAAAAhADHDoo3aAAAABAEAAA8AAAAAAAAAAAAAAAAAuwQAAGRycy9kb3du\ncmV2LnhtbFBLBQYAAAAABAAEAPMAAADCBQAAAAA=\n" };

                V.TextBox textBox2 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "463ACB95", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color8 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize12 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties4.Append(color8);
                paragraphMarkRunProperties4.Append(fontSize12);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript9);

                paragraphProperties4.Append(justification1);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties12 = new RunProperties();
                Color color9 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize13 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "36" };

                runProperties12.Append(color9);
                runProperties12.Append(fontSize13);
                runProperties12.Append(fontSizeComplexScript10);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "School" };
                Tag tag3 = new Tag() { Val = "School" };
                SdtId sdtId4 = new SdtId();
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\' ", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties12);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run7 = new Run();

                RunProperties runProperties13 = new RunProperties();
                Color color10 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize14 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "36" };

                runProperties13.Append(color10);
                runProperties13.Append(fontSize14);
                runProperties13.Append(fontSizeComplexScript11);
                Text text4 = new Text();
                text4.Text = "[School]";

                run7.Append(runProperties13);
                run7.Append(text4);

                sdtContentRun2.Append(run7);

                sdtRun2.Append(sdtProperties4);
                sdtRun2.Append(sdtEndCharProperties4);
                sdtRun2.Append(sdtContentRun2);

                paragraph6.Append(paragraphProperties4);
                paragraph6.Append(sdtRun2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties14 = new RunProperties();
                Color color11 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize15 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "36" };

                runProperties14.Append(color11);
                runProperties14.Append(fontSize15);
                runProperties14.Append(fontSizeComplexScript12);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Course" };
                Tag tag4 = new Tag() { Val = "Course" };
                SdtId sdtId5 = new SdtId();
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns1:category[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties14);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "58A0E1F4", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                Color color12 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize16 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties5.Append(color12);
                paragraphMarkRunProperties5.Append(fontSize16);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript13);

                paragraphProperties5.Append(justification2);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run8 = new Run();

                RunProperties runProperties15 = new RunProperties();
                Color color13 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize17 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "36" };

                runProperties15.Append(color13);
                runProperties15.Append(fontSize17);
                runProperties15.Append(fontSizeComplexScript14);
                Text text5 = new Text();
                text5.Text = "[Course title]";

                run8.Append(runProperties15);
                run8.Append(text5);

                paragraph7.Append(paragraphProperties5);
                paragraph7.Append(run8);

                sdtContentBlock3.Append(paragraph7);

                sdtBlock3.Append(sdtProperties5);
                sdtBlock3.Append(sdtEndCharProperties5);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent2.Append(paragraph6);
                textBoxContent2.Append(sdtBlock3);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap3 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape7.Append(textBox2);
                shape7.Append(textWrap3);

                picture3.Append(shape7);

                run6.Append(runProperties11);
                run6.Append(picture3);

                paragraph2.Append(run1);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00671DD0", RsidRunAdditionDefault = "00A65298", ParagraphId = "1F0817BF", TextId = "7057C8A6" };

                Run run9 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run9.Append(break1);

                paragraph8.Append(run9);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);
                sdtContentBlock1.Append(paragraph8);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;


            }
        }
    }
}
