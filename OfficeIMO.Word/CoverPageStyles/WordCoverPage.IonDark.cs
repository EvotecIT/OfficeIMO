using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageIonDark {
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
                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "5244F80E", TextId = "7BC981DE" };

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "3266355F", TextId = "27C9F0CB" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "1F852C32" };

                V.Group group1 = new V.Group() { Id = "Group 125", Style = "position:absolute;margin-left:0;margin-top:0;width:540pt;height:556.55pt;z-index:-251657216;mso-width-percent:1154;mso-height-percent:670;mso-top-percent:45;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical-relative:page;mso-width-percent:1154;mso-height-percent:670;mso-top-percent:45;mso-width-relative:margin", CoordinateSize = "55613,54044", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                V.Shape shape1 = new V.Shape() { Id = "Freeform 10", Style = "position:absolute;width:55575;height:54044;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", CoordinateSize = "720,700", OptionalString = "_x0000_s1027", FillColor = "#4d5f78 [2994]", Stroked = false, OptionalNumber = 100, Adjustment = "-11796480,,5400", EdgePath = "m,c,644,,644,,644v23,6,62,14,113,21c250,685,476,700,720,644v,-27,,-27,,-27c720,,720,,720,,,,,,,e", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAFyrTUwgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9LasMw\nEN0Xcgcxhe4ayS41xYkSSiDBi0Cp3QMM1sR2Yo2MpcRuTx8VCt3N431nvZ1tL240+s6xhmSpQBDX\nznTcaPiq9s9vIHxANtg7Jg3f5GG7WTysMTdu4k+6laERMYR9jhraEIZcSl+3ZNEv3UAcuZMbLYYI\nx0aaEacYbnuZKpVJix3HhhYH2rVUX8qr1WCS88urU41y5eGn+Kiy49VIr/XT4/y+AhFoDv/iP3dh\n4vw0g99n4gVycwcAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQAFyrTUwgAAANwAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };

                V.Fill fill1 = new V.Fill() { Type = V.FillTypeValues.Gradient, Color2 = "#2a3442 [2018]", Colors = "0 #5d6d85;.5 #485972;1 #334258", Focus = "100%", Rotate = true };
                Ovml.FillExtendedProperties fillExtendedProperties1 = new Ovml.FillExtendedProperties() { Extension = V.ExtensionHandlingBehaviorValues.View, Type = Ovml.FillValues.GradientUnscaled };

                fill1.Append(fillExtendedProperties1);
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Formulas formulas1 = new V.Formulas();
                V.Path path1 = new V.Path() { TextboxRectangle = "0,0,720,700", ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "0,0;0,4972126;872222,5134261;5557520,4972126;5557520,4763667;5557520,0;0,0", ConnectAngles = "0,0,0,0,0,0,0" };

                V.TextBox textBox1 = new V.TextBox() { Inset = "1in,86.4pt,86.4pt,86.4pt" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "7962827D", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "72" };

                paragraphMarkRunProperties1.Append(color1);
                paragraphMarkRunProperties1.Append(fontSize1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "72" };

                runProperties2.Append(color2);
                runProperties2.Append(fontSize2);
                runProperties2.Append(fontSizeComplexScript2);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId();
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
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
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

                textBoxContent1.Append(paragraph3);

                textBox1.Append(textBoxContent1);

                shape1.Append(fill1);
                shape1.Append(stroke1);
                shape1.Append(formulas1);
                shape1.Append(path1);
                shape1.Append(textBox1);

                V.Shape shape2 = new V.Shape() { Id = "Freeform 11", Style = "position:absolute;left:8763;top:47697;width:46850;height:5099;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", CoordinateSize = "607,66", OptionalString = "_x0000_s1028", FillColor = "white [3212]", Stroked = false, EdgePath = "m607,c450,44,300,57,176,57,109,57,49,53,,48,66,58,152,66,251,66,358,66,480,56,607,27,607,,607,,607,e", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDi7/zdwgAAANwAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Ni8Iw\nEL0v+B/CCN7WtB5c6RpFBGEPi2itC70NzdgWm0lpYq3/fiMI3ubxPme5HkwjeupcbVlBPI1AEBdW\n11wqyE67zwUI55E1NpZJwYMcrFejjyUm2t75SH3qSxFC2CWooPK+TaR0RUUG3dS2xIG72M6gD7Ar\npe7wHsJNI2dRNJcGaw4NFba0rai4pjejYLPI/26/1J7z/pDv98f0nMVZrNRkPGy+QXga/Fv8cv/o\nMH/2Bc9nwgVy9Q8AAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDi7/zdwgAAANwAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };
                V.Fill fill2 = new V.Fill() { Opacity = "19789f" };
                V.Path path2 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "4685030,0;1358427,440373;0,370840;1937302,509905;4685030,208598;4685030,0", ConnectAngles = "0,0,0,0,0,0" };

                shape2.Append(fill2);
                shape2.Append(path2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Page };

                group1.Append(lock1);
                group1.Append(shape1);
                group1.Append(shape2);
                group1.Append(textWrap1);

                picture1.Append(group1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                Run run3 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties4.Append(noProof2);

                Picture picture2 = new Picture() { AnchorId = "1E052857" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke2 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path3 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke2);
                shapetype1.Append(path3);

                V.Shape shape3 = new V.Shape() { Id = "Text Box 128", Style = "position:absolute;margin-left:0;margin-top:0;width:453pt;height:11.5pt;z-index:251662336;visibility:visible;mso-wrap-style:square;mso-width-percent:1154;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:bottom;mso-position-vertical-relative:margin;mso-width-percent:1154;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:bottom", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

                V.TextBox textBox2 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "1in,0,86.4pt,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "2DD9E696", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize4 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties2.Append(color4);
                paragraphMarkRunProperties2.Append(fontSize4);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties2.Append(paragraphMarkRunProperties2);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties5 = new RunProperties();
                Caps caps1 = new Caps();
                Color color5 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize5 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "18" };

                runProperties5.Append(caps1);
                runProperties5.Append(color5);
                runProperties5.Append(fontSize5);
                runProperties5.Append(fontSizeComplexScript5);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Company" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId();
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\' ", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties5);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run4 = new Run();

                RunProperties runProperties6 = new RunProperties();
                Caps caps2 = new Caps();
                Color color6 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize6 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "18" };

                runProperties6.Append(caps2);
                runProperties6.Append(color6);
                runProperties6.Append(fontSize6);
                runProperties6.Append(fontSizeComplexScript6);
                Text text2 = new Text();
                text2.Text = "[Company name]";

                run4.Append(runProperties6);
                run4.Append(text2);

                sdtContentRun2.Append(run4);

                sdtRun2.Append(sdtProperties3);
                sdtRun2.Append(sdtEndCharProperties3);
                sdtRun2.Append(sdtContentRun2);

                Run run5 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Caps caps3 = new Caps();
                Color color7 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize7 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "18" };

                runProperties7.Append(caps3);
                runProperties7.Append(color7);
                runProperties7.Append(fontSize7);
                runProperties7.Append(fontSizeComplexScript7);
                Text text3 = new Text();
                text3.Text = " ";

                run5.Append(runProperties7);
                run5.Append(text3);

                Run run6 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Color color8 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize8 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "18" };

                runProperties8.Append(color8);
                runProperties8.Append(fontSize8);
                runProperties8.Append(fontSizeComplexScript8);
                Text text4 = new Text();
                text4.Text = "| ";

                run6.Append(runProperties8);
                run6.Append(text4);

                SdtRun sdtRun3 = new SdtRun();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                Color color9 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize9 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "18" };

                runProperties9.Append(color9);
                runProperties9.Append(fontSize9);
                runProperties9.Append(fontSizeComplexScript9);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Address" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId();
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:CompanyAddress[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties9);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun3 = new SdtContentRun();

                Run run7 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Color color10 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Text1, ThemeTint = "80" };
                FontSize fontSize10 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "18" };

                runProperties10.Append(color10);
                runProperties10.Append(fontSize10);
                runProperties10.Append(fontSizeComplexScript10);
                Text text5 = new Text();
                text5.Text = "[Company address]";

                run7.Append(runProperties10);
                run7.Append(text5);

                sdtContentRun3.Append(run7);

                sdtRun3.Append(sdtProperties4);
                sdtRun3.Append(sdtEndCharProperties4);
                sdtRun3.Append(sdtContentRun3);

                paragraph4.Append(paragraphProperties2);
                paragraph4.Append(sdtRun2);
                paragraph4.Append(run5);
                paragraph4.Append(run6);
                paragraph4.Append(sdtRun3);

                textBoxContent2.Append(paragraph4);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape3.Append(textBox2);
                shape3.Append(textWrap2);

                picture2.Append(shapetype1);
                picture2.Append(shape3);

                run3.Append(runProperties4);
                run3.Append(picture2);

                Run run8 = new Run();

                RunProperties runProperties11 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties11.Append(noProof3);

                Picture picture3 = new Picture() { AnchorId = "5A33438E" };

                V.Shape shape4 = new V.Shape() { Id = "Text Box 129", Style = "position:absolute;margin-left:0;margin-top:0;width:453pt;height:38.15pt;z-index:251661312;visibility:visible;mso-wrap-style:square;mso-width-percent:1154;mso-height-percent:0;mso-top-percent:790;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:1154;mso-height-percent:0;mso-top-percent:790;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", OptionalString = "_x0000_s1030", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\n" };

                V.TextBox textBox3 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "1in,0,86.4pt,0" };

                TextBoxContent textBoxContent3 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties12 = new RunProperties();
                Caps caps4 = new Caps();
                Color color11 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize11 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

                runProperties12.Append(caps4);
                runProperties12.Append(color11);
                runProperties12.Append(fontSize11);
                runProperties12.Append(fontSizeComplexScript11);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Subtitle" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId();
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties12);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "2A0EF94F", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "40", After = "40" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Caps caps5 = new Caps();
                Color color12 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize12 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties3.Append(caps5);
                paragraphMarkRunProperties3.Append(color12);
                paragraphMarkRunProperties3.Append(fontSize12);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript12);

                paragraphProperties3.Append(spacingBetweenLines1);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run9 = new Run();

                RunProperties runProperties13 = new RunProperties();
                Caps caps6 = new Caps();
                Color color13 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize13 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

                runProperties13.Append(caps6);
                runProperties13.Append(color13);
                runProperties13.Append(fontSize13);
                runProperties13.Append(fontSizeComplexScript13);
                Text text6 = new Text();
                text6.Text = "[Document subtitle]";

                run9.Append(runProperties13);
                run9.Append(text6);

                paragraph5.Append(paragraphProperties3);
                paragraph5.Append(run9);

                sdtContentBlock2.Append(paragraph5);

                sdtBlock2.Append(sdtProperties5);
                sdtBlock2.Append(sdtEndCharProperties5);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties6 = new SdtProperties();

                RunProperties runProperties14 = new RunProperties();
                Caps caps7 = new Caps();
                Color color14 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize14 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

                runProperties14.Append(caps7);
                runProperties14.Append(color14);
                runProperties14.Append(fontSize14);
                runProperties14.Append(fontSizeComplexScript14);
                SdtAlias sdtAlias5 = new SdtAlias() { Val = "Author" };
                Tag tag5 = new Tag() { Val = "" };
                SdtId sdtId6 = new SdtId();
                ShowingPlaceholder showingPlaceholder5 = new ShowingPlaceholder();
                DataBinding dataBinding5 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText5 = new SdtContentText();

                sdtProperties6.Append(runProperties14);
                sdtProperties6.Append(sdtAlias5);
                sdtProperties6.Append(tag5);
                sdtProperties6.Append(sdtId6);
                sdtProperties6.Append(showingPlaceholder5);
                sdtProperties6.Append(dataBinding5);
                sdtProperties6.Append(sdtContentText5);
                SdtEndCharProperties sdtEndCharProperties6 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "1B7F11D6", TextId = "6066EBB9" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "40", After = "40" };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Caps caps8 = new Caps();
                Color color15 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize15 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties4.Append(caps8);
                paragraphMarkRunProperties4.Append(color15);
                paragraphMarkRunProperties4.Append(fontSize15);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript15);

                paragraphProperties4.Append(spacingBetweenLines2);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run10 = new Run();

                RunProperties runProperties15 = new RunProperties();
                Caps caps9 = new Caps();
                Color color16 = new Color() { Val = "5B9BD5", ThemeColor = ThemeColorValues.Accent5 };
                FontSize fontSize16 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

                runProperties15.Append(caps9);
                runProperties15.Append(color16);
                runProperties15.Append(fontSize16);
                runProperties15.Append(fontSizeComplexScript16);
                Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text7.Text = "     ";

                run10.Append(runProperties15);
                run10.Append(text7);

                paragraph6.Append(paragraphProperties4);
                paragraph6.Append(run10);

                sdtContentBlock3.Append(paragraph6);

                sdtBlock3.Append(sdtProperties6);
                sdtBlock3.Append(sdtEndCharProperties6);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent3.Append(sdtBlock2);
                textBoxContent3.Append(sdtBlock3);

                textBox3.Append(textBoxContent3);
                Wvml.TextWrap textWrap3 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape4.Append(textBox3);
                shape4.Append(textWrap3);

                picture3.Append(shape4);

                run8.Append(runProperties11);
                run8.Append(picture3);

                Run run11 = new Run();

                RunProperties runProperties16 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties16.Append(noProof4);

                Picture picture4 = new Picture() { AnchorId = "7181A95C" };

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 130", Style = "position:absolute;margin-left:-8.8pt;margin-top:0;width:46.8pt;height:77.75pt;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-width-percent:76;mso-height-percent:98;mso-top-percent:23;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:right;mso-position-horizontal-relative:margin;mso-position-vertical-relative:page;mso-width-percent:76;mso-height-percent:98;mso-top-percent:23;mso-width-relative:page;mso-height-relative:page;v-text-anchor:bottom", OptionalString = "_x0000_s1031", FillColor = "#4472c4 [3204]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAOUFAAAAAA==\n"));
                Ovml.Lock lock2 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

                V.TextBox textBox4 = new V.TextBox() { Inset = "3.6pt,,3.6pt" };

                TextBoxContent textBoxContent4 = new TextBoxContent();

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties7 = new SdtProperties();

                RunProperties runProperties17 = new RunProperties();
                Color color17 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize17 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

                runProperties17.Append(color17);
                runProperties17.Append(fontSize17);
                runProperties17.Append(fontSizeComplexScript17);
                SdtAlias sdtAlias6 = new SdtAlias() { Val = "Year" };
                Tag tag6 = new Tag() { Val = "" };
                SdtId sdtId7 = new SdtId();
                ShowingPlaceholder showingPlaceholder6 = new ShowingPlaceholder();
                DataBinding dataBinding6 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };

                SdtContentDate sdtContentDate1 = new SdtContentDate() { FullDate = System.Xml.XmlConvert.ToDateTime("2012-03-16T00:00:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind) };
                DateFormat dateFormat1 = new DateFormat() { Val = "yyyy" };
                LanguageId languageId1 = new LanguageId() { Val = "en-US" };
                SdtDateMappingType sdtDateMappingType1 = new SdtDateMappingType() { Val = DateFormatValues.DateTime };
                Calendar calendar1 = new Calendar() { Val = CalendarValues.Gregorian };

                sdtContentDate1.Append(dateFormat1);
                sdtContentDate1.Append(languageId1);
                sdtContentDate1.Append(sdtDateMappingType1);
                sdtContentDate1.Append(calendar1);

                sdtProperties7.Append(runProperties17);
                sdtProperties7.Append(sdtAlias6);
                sdtProperties7.Append(tag6);
                sdtProperties7.Append(sdtId7);
                sdtProperties7.Append(showingPlaceholder6);
                sdtProperties7.Append(dataBinding6);
                sdtProperties7.Append(sdtContentDate1);
                SdtEndCharProperties sdtEndCharProperties7 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00A6334D", RsidRunAdditionDefault = "005279DA", ParagraphId = "34FF1EFB", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                Color color18 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize18 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties5.Append(color18);
                paragraphMarkRunProperties5.Append(fontSize18);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript18);

                paragraphProperties5.Append(justification1);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run12 = new Run();

                RunProperties runProperties18 = new RunProperties();
                Color color19 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize19 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

                runProperties18.Append(color19);
                runProperties18.Append(fontSize19);
                runProperties18.Append(fontSizeComplexScript19);
                Text text8 = new Text();
                text8.Text = "[Year]";

                run12.Append(runProperties18);
                run12.Append(text8);

                paragraph7.Append(paragraphProperties5);
                paragraph7.Append(run12);

                sdtContentBlock4.Append(paragraph7);

                sdtBlock4.Append(sdtProperties7);
                sdtBlock4.Append(sdtEndCharProperties7);
                sdtBlock4.Append(sdtContentBlock4);

                textBoxContent4.Append(sdtBlock4);

                textBox4.Append(textBoxContent4);
                Wvml.TextWrap textWrap4 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle1.Append(lock2);
                rectangle1.Append(textBox4);
                rectangle1.Append(textWrap4);

                picture4.Append(rectangle1);

                run11.Append(runProperties16);
                run11.Append(picture4);

                Run run13 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run13.Append(break1);

                paragraph2.Append(run1);
                paragraph2.Append(run3);
                paragraph2.Append(run8);
                paragraph2.Append(run11);
                paragraph2.Append(run13);

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
