using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageViewMaster {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1338198481 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00545D80", RsidRunAdditionDefault = "00AC070C", ParagraphId = "7EC01ED6", TextId = "6F1C331E" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties1.Append(noProof1);
                runProperties1.Append(color1);

                Picture picture1 = new Picture() { AnchorId = "02630B32" };

                V.Group group1 = new V.Group() { Id = "Group 11", Style = "position:absolute;margin-left:0;margin-top:0;width:540pt;height:10in;z-index:251659264;mso-width-percent:882;mso-height-percent:909;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:882;mso-height-percent:909", CoordinateSize = "68580,91440", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQCdoio4vgMAALoOAAAOAAAAZHJzL2Uyb0RvYy54bWzsV9tu4zYQfS/QfyD43shybCUrRFmk2SYo\nkO4GmxT7TFOUJZQiWZKOlH59h0NJWTtuN3WRboH2ReJlbjyaORydve1bSR6EdY1WBU2PZpQIxXXZ\nqHVBf76/+u6UEueZKpnUShT0UTj69vzbb846k4u5rrUshSVgRLm8MwWtvTd5kjhei5a5I22Egs1K\n25Z5mNp1UlrWgfVWJvPZLEs6bUtjNRfOweq7uEnP0X5VCe4/VJUTnsiCQmwenxafq/BMzs9YvrbM\n1A0fwmAHRNGyRoHTydQ75hnZ2OaZqbbhVjtd+SOu20RXVcMFngFOk852TnNt9cbgWdZ5tzYTTADt\nDk4Hm+XvH66tuTO3FpDozBqwwFk4S1/ZNrwhStIjZI8TZKL3hMNidro8nc0AWQ57b9LFIkwQVF4D\n8s/0eP3DFzST0XGyFU5nIEHcEwbu72FwVzMjEFqXAwa3ljRlQY+PKVGshTz9CJnD1FoKAmsIDcpN\nQLncAWZ7UJrPT7OAxx6osvkbgOcZVNOBWW6s89dCtyQMCmohCMwq9nDjPEQBoqNIcO20bMqrRkqc\nhJIRl9KSBwbJ7vs0xA0aW1JSBVmlg1bcDisA9XggHPlHKYKcVB9FBcjAh55jIFiXT04Y50L5NG7V\nrBTR9xKSALMgeB/DwljQYLBcgf/J9mBglIxGRtsxykE+qAos60l59meBReVJAz1r5SfltlHa7jMg\n4VSD5yg/ghShCSj5ftWDSBiudPkIKWR15Bdn+FUDX/CGOX/LLBAKfHUgSf8BHpXUXUH1MKKk1va3\nfetBHnIcdinpgKAK6n7dMCsokT8qyP7F8gQIECgNZ7H0KLFbsxXO5tkyPclAVG3aSw25kQInG45D\nWLVejsPK6vYTEOpFcA1bTHEIoKCrcXjpI3cCIXNxcYFCQGSG+Rt1Z3gwHTAOSXrff2LWDJnsgS/e\n67HoWL6T0FE2aCp9sfG6ajDbn6Ad0AcCiIC/PhMs9jDB4i8xAeD3nARGetihy9fjgACq3LQ/6XKn\nNnE5UDTSxR+X7P/08Tr08TU4A1lipAwkkC3GoCQSxrBzMF1wb/97hLEcCeM+9Ebf654cL3f4gvge\n1gNJDvfGl3uIkzSbQ4MVFODmG1unzzuJdJmliwxdHc4iU0MQ7nwC91N2vIwX67QDxmPzEMthaEiw\nXcBOCEd7GocX3M/7u4IXKP7TXUH5y4u6Auwlp4/8NQp9uzkIl/yevgCWDy7xf1VHgH8K8IOELebw\nMxf+wD6fYwfx9Mt5/jsAAAD//wMAUEsDBBQABgAIAAAAIQCQ+IEL2gAAAAcBAAAPAAAAZHJzL2Rv\nd25yZXYueG1sTI9BT8MwDIXvSPyHyEjcWMI0TVNpOqFJ4wSHrbtw8xLTVmucqsm28u/xuMDFek/P\nev5crqfQqwuNqYts4XlmQBG76DtuLBzq7dMKVMrIHvvIZOGbEqyr+7sSCx+vvKPLPjdKSjgVaKHN\neSi0Tq6lgGkWB2LJvuIYMIsdG+1HvEp56PXcmKUO2LFcaHGgTUvutD8HC6fdR6LNtm4OLrhuOb2/\nzT/rYO3jw/T6AirTlP+W4YYv6FAJ0zGe2SfVW5BH8u+8ZWZlxB9FLRaidFXq//zVDwAAAP//AwBQ\nSwECLQAUAAYACAAAACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlw\nZXNdLnhtbFBLAQItABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAAAC8BAABfcmVs\ncy8ucmVsc1BLAQItABQABgAIAAAAIQCdoio4vgMAALoOAAAOAAAAAAAAAAAAAAAAAC4CAABkcnMv\nZTJvRG9jLnhtbFBLAQItABQABgAIAAAAIQCQ+IEL2gAAAAcBAAAPAAAAAAAAAAAAAAAAABgGAABk\ncnMvZG93bnJldi54bWxQSwUGAAAAAAQABADzAAAAHwcAAAAA\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 33", Style = "position:absolute;left:2286;width:66294;height:91440;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", OptionalString = "_x0000_s1027", FillColor = "black [3213]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQA4mHyFwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Pi8Iw\nFMTvgt8hPGFvmqoopdtURBE8LIh/8Pxs3rbdbV5KE2v3228EweMwM79h0lVvatFR6yrLCqaTCARx\nbnXFhYLLeTeOQTiPrLG2TAr+yMEqGw5STLR98JG6ky9EgLBLUEHpfZNI6fKSDLqJbYiD921bgz7I\ntpC6xUeAm1rOomgpDVYcFkpsaFNS/nu6GwV9vO0W3F3vx/XtwGa7+7r95LFSH6N+/QnCU+/f4Vd7\nrxXM5/D8En6AzP4BAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAOJh8hcMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n"));

                V.TextBox textBox1 = new V.TextBox() { Inset = "36pt,1in,1in,208.8pt" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "84" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "84" };

                runProperties2.Append(runFonts1);
                runProperties2.Append(color2);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId() { Val = -960264625 };
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

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00545D80", RsidRunAdditionDefault = "00AC070C", ParagraphId = "0CB89AFA", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "120" };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "84" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "84" };

                paragraphMarkRunProperties1.Append(runFonts2);
                paragraphMarkRunProperties1.Append(color3);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(spacingBetweenLines1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "84" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "84" };

                runProperties3.Append(runFonts3);
                runProperties3.Append(color4);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "[Document title]";

                run2.Append(runProperties3);
                run2.Append(text1);

                paragraph2.Append(paragraphProperties1);
                paragraph2.Append(run2);

                sdtContentBlock2.Append(paragraph2);

                sdtBlock2.Append(sdtProperties2);
                sdtBlock2.Append(sdtEndCharProperties2);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties4.Append(color5);
                runProperties4.Append(fontSize4);
                runProperties4.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Subtitle" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId() { Val = 1611937615 };
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

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00545D80", RsidRunAdditionDefault = "00AC070C", ParagraphId = "211E26CA", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties2.Append(color6);
                paragraphMarkRunProperties2.Append(fontSize5);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript5);

                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Color color7 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(color7);
                runProperties5.Append(fontSize6);
                runProperties5.Append(fontSizeComplexScript6);
                Text text2 = new Text();
                text2.Text = "[Document subtitle]";

                run3.Append(runProperties5);
                run3.Append(text2);

                paragraph3.Append(paragraphProperties2);
                paragraph3.Append(run3);

                sdtContentBlock3.Append(paragraph3);

                sdtBlock3.Append(sdtProperties3);
                sdtBlock3.Append(sdtEndCharProperties3);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent1.Append(sdtBlock2);
                textBoxContent1.Append(sdtBlock3);

                textBox1.Append(textBoxContent1);

                rectangle1.Append(textBox1);

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 34", Style = "position:absolute;width:2286;height:91440;visibility:visible;mso-wrap-style:square;v-text-anchor:middle", OptionalString = "_x0000_s1028", FillColor = "gray [1629]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCre5UexAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvhf6H5RV6Ed1Yi0jqKhpQBCloFOnxkX0mwezbmF1N/PfdgtDjMDPfMNN5Zypxp8aVlhUMBxEI\n4szqknMFx8OqPwHhPLLGyjIpeJCD+ez1ZYqxti3v6Z76XAQIuxgVFN7XsZQuK8igG9iaOHhn2xj0\nQTa51A22AW4q+RFFY2mw5LBQYE1JQdklvRkFvZ/TNll6/X25JjWd7W7dLlOj1Ptbt/gC4anz/+Fn\ne6MVjD7h70v4AXL2CwAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAKt7lR7EAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 35", Style = "position:absolute;left:2286;top:71628;width:66294;height:15614;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAn8efKxAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Pa8JA\nFMTvgt9heUJvurHVotFV2kJLqV78c/H2yD6TYPZtmn016bd3hUKPw8z8hlmuO1epKzWh9GxgPEpA\nEWfelpwbOB7ehzNQQZAtVp7JwC8FWK/6vSWm1re8o+techUhHFI0UIjUqdYhK8hhGPmaOHpn3ziU\nKJtc2wbbCHeVfkySZ+2w5LhQYE1vBWWX/Y8zcPrefoXJpp27V57iLpnIx3YsxjwMupcFKKFO/sN/\n7U9r4GkK9y/xB+jVDQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhACfx58rEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n" };

                V.TextBox textBox2 = new V.TextBox() { Inset = "36pt,0,1in,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties6 = new RunProperties();
                Color color8 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize7 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "32" };

                runProperties6.Append(color8);
                runProperties6.Append(fontSize7);
                runProperties6.Append(fontSizeComplexScript7);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Author" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId() { Val = -315646564 };
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

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00545D80", RsidRunAdditionDefault = "00AC070C", ParagraphId = "7892C93C", TextId = "59FD6EE1" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color9 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize8 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties3.Append(color9);
                paragraphMarkRunProperties3.Append(fontSize8);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript8);

                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run4 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Color color10 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize9 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "32" };

                runProperties7.Append(color10);
                runProperties7.Append(fontSize9);
                runProperties7.Append(fontSizeComplexScript9);
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = "     ";

                run4.Append(runProperties7);
                run4.Append(text3);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(run4);

                sdtContentBlock4.Append(paragraph4);

                sdtBlock4.Append(sdtProperties4);
                sdtBlock4.Append(sdtEndCharProperties4);
                sdtBlock4.Append(sdtContentBlock4);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00545D80", RsidRunAdditionDefault = "00AC070C", ParagraphId = "323F0579", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color11 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize10 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties4.Append(color11);
                paragraphMarkRunProperties4.Append(fontSize10);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript10);

                paragraphProperties4.Append(paragraphMarkRunProperties4);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties8 = new RunProperties();
                Caps caps1 = new Caps();
                Color color12 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize11 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "18" };

                runProperties8.Append(caps1);
                runProperties8.Append(color12);
                runProperties8.Append(fontSize11);
                runProperties8.Append(fontSizeComplexScript11);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Company" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId() { Val = -775099975 };
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\' ", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties8);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run5 = new Run();

                RunProperties runProperties9 = new RunProperties();
                Caps caps2 = new Caps();
                Color color13 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize12 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "18" };

                runProperties9.Append(caps2);
                runProperties9.Append(color13);
                runProperties9.Append(fontSize12);
                runProperties9.Append(fontSizeComplexScript12);
                Text text4 = new Text();
                text4.Text = "[Company name]";

                run5.Append(runProperties9);
                run5.Append(text4);

                sdtContentRun1.Append(run5);

                sdtRun1.Append(sdtProperties5);
                sdtRun1.Append(sdtEndCharProperties5);
                sdtRun1.Append(sdtContentRun1);

                Run run6 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Color color14 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize13 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "18" };

                runProperties10.Append(color14);
                runProperties10.Append(fontSize13);
                runProperties10.Append(fontSizeComplexScript13);
                Text text5 = new Text();
                text5.Text = "  ";

                run6.Append(runProperties10);
                run6.Append(text5);

                SdtRun sdtRun2 = new SdtRun();

                SdtProperties sdtProperties6 = new SdtProperties();

                RunProperties runProperties11 = new RunProperties();
                Color color15 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize14 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "18" };

                runProperties11.Append(color15);
                runProperties11.Append(fontSize14);
                runProperties11.Append(fontSizeComplexScript14);
                SdtAlias sdtAlias5 = new SdtAlias() { Val = "Address" };
                Tag tag5 = new Tag() { Val = "" };
                SdtId sdtId6 = new SdtId() { Val = -669564449 };
                ShowingPlaceholder showingPlaceholder5 = new ShowingPlaceholder();
                DataBinding dataBinding5 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:CompanyAddress[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText5 = new SdtContentText();

                sdtProperties6.Append(runProperties11);
                sdtProperties6.Append(sdtAlias5);
                sdtProperties6.Append(tag5);
                sdtProperties6.Append(sdtId6);
                sdtProperties6.Append(showingPlaceholder5);
                sdtProperties6.Append(dataBinding5);
                sdtProperties6.Append(sdtContentText5);
                SdtEndCharProperties sdtEndCharProperties6 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun2 = new SdtContentRun();

                Run run7 = new Run();

                RunProperties runProperties12 = new RunProperties();
                Color color16 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize15 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "18" };

                runProperties12.Append(color16);
                runProperties12.Append(fontSize15);
                runProperties12.Append(fontSizeComplexScript15);
                Text text6 = new Text();
                text6.Text = "[Company address]";

                run7.Append(runProperties12);
                run7.Append(text6);

                sdtContentRun2.Append(run7);

                sdtRun2.Append(sdtProperties6);
                sdtRun2.Append(sdtEndCharProperties6);
                sdtRun2.Append(sdtContentRun2);

                paragraph5.Append(paragraphProperties4);
                paragraph5.Append(sdtRun1);
                paragraph5.Append(run6);
                paragraph5.Append(sdtRun2);

                textBoxContent2.Append(sdtBlock4);
                textBoxContent2.Append(paragraph5);

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

                Run run8 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run8.Append(break1);

                paragraph1.Append(run1);
                paragraph1.Append(run8);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;


            }
        }
    }
}
