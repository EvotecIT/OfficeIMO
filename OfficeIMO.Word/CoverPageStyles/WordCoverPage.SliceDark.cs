using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageSliceDark {
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

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "008167E7", RsidRunAdditionDefault = "006B6A4E", ParagraphId = "130A51B8", TextId = "5FBB28D0" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "497650A3" };

                V.Group group1 = new V.Group() { Id = "Group 48", Style = "position:absolute;margin-left:0;margin-top:0;width:540pt;height:10in;z-index:-251655168;mso-width-percent:882;mso-height-percent:909;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:882;mso-height-percent:909", CoordinateSize = "68580,91440", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAAAA\n"));

                V.Group group2 = new V.Group() { Id = "Group 49", Style = "position:absolute;width:68580;height:91440", CoordinateSize = "68580,91440", OptionalString = "_x0000_s1027" };
                group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBvtESaxgAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9ba8JA\nFITfC/6H5Qh9q5vYVjRmFRFb+iCCFxDfDtmTC2bPhuw2if++Wyj0cZiZb5h0PZhadNS6yrKCeBKB\nIM6srrhQcDl/vMxBOI+ssbZMCh7kYL0aPaWYaNvzkbqTL0SAsEtQQel9k0jpspIMuoltiIOX29ag\nD7ItpG6xD3BTy2kUzaTBisNCiQ1tS8rup2+j4LPHfvMa77r9Pd8+buf3w3Ufk1LP42GzBOFp8P/h\nv/aXVvC2gN8v4QfI1Q8AAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAA\nAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAA\nCwAAAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAb7REmsYAAADbAAAA\nDwAAAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPoCAAAAAA==\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 54", Style = "position:absolute;width:68580;height:91440;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", FillColor = "#485870 [3122]", Stroked = false, StrokeWeight = "1pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCGyzIlxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx\nFITvBf9DeIK3mq22pWyNooK0eBFtBXt73Tw3wc3LuknX9d+bQqHHYWa+YSazzlWipSZYzwoehhkI\n4sJry6WCz4/V/QuIEJE1Vp5JwZUCzKa9uwnm2l94S+0uliJBOOSowMRY51KGwpDDMPQ1cfKOvnEY\nk2xKqRu8JLir5CjLnqVDy2nBYE1LQ8Vp9+MUbPft+GA26zdr7XjxffVy/XU+KjXod/NXEJG6+B/+\na79rBU+P8Psl/QA5vQEAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQCGyzIlxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n"));
                V.Fill fill1 = new V.Fill() { Type = V.FillTypeValues.Gradient, Color2 = "#3d4b5f [2882]", Colors = "0 #88acbb;6554f #88acbb", Angle = 348M, Focus = "100%" };

                V.TextBox textBox1 = new V.TextBox() { Inset = "54pt,54pt,1in,5in" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "008167E7", RsidRunAdditionDefault = "006B6A4E", ParagraphId = "60E49FBC", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "48" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "48" };

                paragraphMarkRunProperties1.Append(color1);
                paragraphMarkRunProperties1.Append(fontSize1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties1.Append(paragraphMarkRunProperties1);

                paragraph2.Append(paragraphProperties1);

                textBoxContent1.Append(paragraph2);

                textBox1.Append(textBoxContent1);

                rectangle1.Append(fill1);
                rectangle1.Append(textBox1);

                V.Group group3 = new V.Group() { Id = "Group 2", Style = "position:absolute;left:25241;width:43291;height:44910", CoordinateSize = "43291,44910", OptionalString = "_x0000_s1029" };
                group3.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBrINhCwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Bi8Iw\nFITvC/sfwlvwtqZVKkvXKCKreBBBXRBvj+bZFpuX0sS2/nsjCB6HmfmGmc57U4mWGldaVhAPIxDE\nmdUl5wr+j6vvHxDOI2usLJOCOzmYzz4/pphq2/Ge2oPPRYCwS1FB4X2dSumyggy6oa2Jg3exjUEf\nZJNL3WAX4KaSoyiaSIMlh4UCa1oWlF0PN6Ng3WG3GMd/7fZ6Wd7Px2R32sak1OCrX/yC8NT7d/jV\n3mgFSQLPL+EHyNkDAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAayDYQsMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n"));

                V.Shape shape1 = new V.Shape() { Id = "Freeform 56", Style = "position:absolute;left:15017;width:28274;height:28352;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "1781,1786", OptionalString = "_x0000_s1030", Filled = false, Stroked = false, EdgePath = "m4,1786l,1782,1776,r5,5l4,1786xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCTDSSBwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx\nFITvQv9DeIXeNFuhYlej2MK23mq3xfNj89wNbl62SVzXf28KgsdhZr5hluvBtqInH4xjBc+TDARx\n5bThWsHvTzGegwgRWWPrmBRcKMB69TBaYq7dmb+pL2MtEoRDjgqaGLtcylA1ZDFMXEecvIPzFmOS\nvpba4znBbSunWTaTFg2nhQY7em+oOpYnq6B/88NXdPttUZjdq+z1h/n73Cv19DhsFiAiDfEevrW3\nWsHLDP6/pB8gV1cAAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAkw0kgcMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };
                V.Path path1 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "6350,2835275;0,2828925;2819400,0;2827338,7938;6350,2835275", ConnectAngles = "0,0,0,0,0" };

                shape1.Append(path1);

                V.Shape shape2 = new V.Shape() { Id = "Freeform 57", Style = "position:absolute;left:7826;top:2270;width:35465;height:35464;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2234,2234", OptionalString = "_x0000_s1031", Filled = false, Stroked = false, EdgePath = "m5,2234l,2229,2229,r5,5l5,2234xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAcaGAJxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/RasJA\nFETfC/7DcoW+NRsttTV1FRHFPoil0Q+4zV6TYPZuzG5i2q93hUIfh5k5w8wWvalER40rLSsYRTEI\n4szqknMFx8Pm6Q2E88gaK8uk4IccLOaDhxkm2l75i7rU5yJA2CWooPC+TqR0WUEGXWRr4uCdbGPQ\nB9nkUjd4DXBTyXEcT6TBksNCgTWtCsrOaWsU9L/tdve5HtW7STV99t/yspruUanHYb98B+Gp9//h\nv/aHVvDyCvcv4QfI+Q0AAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQAcaGAJxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n" };
                V.Path path2 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "7938,3546475;0,3538538;3538538,0;3546475,7938;7938,3546475", ConnectAngles = "0,0,0,0,0" };

                shape2.Append(path2);

                V.Shape shape3 = new V.Shape() { Id = "Freeform 58", Style = "position:absolute;left:8413;top:1095;width:34878;height:34877;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2197,2197", OptionalString = "_x0000_s1032", Filled = false, Stroked = false, EdgePath = "m9,2197l,2193,2188,r9,10l9,2197xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDUx4njwgAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Na8JA\nEL0L/odlCr2I2ViwhugmSCFtr1VL8TZmxyQ0O5tmt0n8991DwePjfe/yybRioN41lhWsohgEcWl1\nw5WC07FYJiCcR9bYWiYFN3KQZ/PZDlNtR/6g4eArEULYpaig9r5LpXRlTQZdZDviwF1tb9AH2FdS\n9ziGcNPKpzh+lgYbDg01dvRSU/l9+DUKEnceN0f8eR28vK6axeWz+HorlHp8mPZbEJ4mfxf/u9+1\ngnUYG76EHyCzPwAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDUx4njwgAAANsAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };
                V.Path path3 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "14288,3487738;0,3481388;3473450,0;3487738,15875;14288,3487738", ConnectAngles = "0,0,0,0,0" };

                shape3.Append(path3);

                V.Shape shape4 = new V.Shape() { Id = "Freeform 59", Style = "position:absolute;left:12160;top:4984;width:31131;height:31211;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "1961,1966", OptionalString = "_x0000_s1033", Filled = false, Stroked = false, EdgePath = "m9,1966l,1957,1952,r9,9l9,1966xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQANjfa1wwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/BagIx\nEIbvgu8QRuhNs0or7moUaVGk0INa6HXcTDdLN5Mlie769k2h4HH45//mm9Wmt424kQ+1YwXTSQaC\nuHS65krB53k3XoAIEVlj45gU3CnAZj0crLDQruMj3U6xEgnCoUAFJsa2kDKUhiyGiWuJU/btvMWY\nRl9J7bFLcNvIWZbNpcWa0wWDLb0aKn9OV5s0vmZv+2cjL8lqnn0c97l/73Klnkb9dgkiUh8fy//t\ng1bwksPfLwkAcv0LAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEADY32tcMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };
                V.Path path4 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "14288,3121025;0,3106738;3098800,0;3113088,14288;14288,3121025", ConnectAngles = "0,0,0,0,0" };

                shape4.Append(path4);

                V.Shape shape5 = new V.Shape() { Id = "Freeform 60", Style = "position:absolute;top:1539;width:43291;height:43371;visibility:visible;mso-wrap-style:square;v-text-anchor:top", CoordinateSize = "2727,2732", OptionalString = "_x0000_s1034", Filled = false, Stroked = false, EdgePath = "m,2732r,-4l2722,r5,5l,2732xe", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCvi0/huwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9LCsIw\nEN0L3iGM4E5TXZRSjaUIgi79HGBopm2wmZQmavX0ZiG4fLz/thhtJ540eONYwWqZgCCunDbcKLhd\nD4sMhA/IGjvHpOBNHorddLLFXLsXn+l5CY2IIexzVNCG0OdS+qoli37peuLI1W6wGCIcGqkHfMVw\n28l1kqTSouHY0GJP+5aq++VhFSRmferOaW20rLP7zZyyY/mplJrPxnIDItAY/uKf+6gVpHF9/BJ/\ngNx9AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAAAABb\nQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAAAAAA\nAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAK+LT+G7AAAA2wAAAA8AAAAAAAAAAAAA\nAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADvAgAAAAA=\n" };
                V.Path path5 = new V.Path() { ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "0,4337050;0,4330700;4321175,0;4329113,7938;0,4337050", ConnectAngles = "0,0,0,0,0" };

                shape5.Append(path5);

                group3.Append(shape1);
                group3.Append(shape2);
                group3.Append(shape3);
                group3.Append(shape4);
                group3.Append(shape5);

                group2.Append(rectangle1);
                group2.Append(group3);

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path6 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path6);

                V.Shape shape6 = new V.Shape() { Id = "Text Box 61", Style = "position:absolute;left:95;top:48387;width:68434;height:37897;visibility:visible;mso-wrap-style:square;v-text-anchor:bottom", OptionalString = "_x0000_s1035", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCY+K1FwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/NasMw\nEITvhb6D2EJvtZweQnGjhJCQOsfmr/S4WFtLxFo5lmq7b18FAjkOM/MNM1uMrhE9dcF6VjDJchDE\nldeWawXHw+blDUSIyBobz6TgjwIs5o8PMyy0H3hH/T7WIkE4FKjAxNgWUobKkMOQ+ZY4eT++cxiT\n7GqpOxwS3DXyNc+n0qHltGCwpZWh6rz/dQoG7q0tZbP+kp/56bv8MNtLuVPq+WlcvoOINMZ7+Nbe\nagXTCVy/pB8g5/8AAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAmPitRcMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };

                V.TextBox textBox2 = new V.TextBox() { Inset = "54pt,0,1in,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps1 = new Caps();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "64" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "64" };

                runProperties2.Append(runFonts1);
                runProperties2.Append(caps1);
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

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "008167E7", RsidRunAdditionDefault = "006B6A4E", ParagraphId = "124D4103", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps2 = new Caps();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "64" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "64" };

                paragraphMarkRunProperties2.Append(runFonts2);
                paragraphMarkRunProperties2.Append(caps2);
                paragraphMarkRunProperties2.Append(color3);
                paragraphMarkRunProperties2.Append(fontSize3);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Caps caps3 = new Caps();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize4 = new FontSize() { Val = "64" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "64" };

                runProperties3.Append(runFonts3);
                runProperties3.Append(caps3);
                runProperties3.Append(color4);
                runProperties3.Append(fontSize4);
                runProperties3.Append(fontSizeComplexScript4);
                Text text1 = new Text();
                text1.Text = "[Document title]";

                run2.Append(runProperties3);
                run2.Append(text1);

                paragraph3.Append(paragraphProperties2);
                paragraph3.Append(run2);

                sdtContentBlock2.Append(paragraph3);

                sdtBlock2.Append(sdtProperties2);
                sdtBlock2.Append(sdtEndCharProperties2);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Color color5 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize5 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "36" };

                runProperties4.Append(color5);
                runProperties4.Append(fontSize5);
                runProperties4.Append(fontSizeComplexScript5);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Subtitle" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId();
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

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "008167E7", RsidRunAdditionDefault = "006B6A4E", ParagraphId = "38C3D613", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "120" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color6 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize6 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "36" };

                paragraphMarkRunProperties3.Append(color6);
                paragraphMarkRunProperties3.Append(fontSize6);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript6);

                paragraphProperties3.Append(spacingBetweenLines1);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Color color7 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize7 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "36" };

                runProperties5.Append(color7);
                runProperties5.Append(fontSize7);
                runProperties5.Append(fontSizeComplexScript7);
                Text text2 = new Text();
                text2.Text = "[Document subtitle]";

                run3.Append(runProperties5);
                run3.Append(text2);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(run3);

                sdtContentBlock3.Append(paragraph4);

                sdtBlock3.Append(sdtProperties3);
                sdtBlock3.Append(sdtEndCharProperties3);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent2.Append(sdtBlock2);
                textBoxContent2.Append(sdtBlock3);

                textBox2.Append(textBoxContent2);

                shape6.Append(textBox2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                group1.Append(group2);
                group1.Append(shapetype1);
                group1.Append(shape6);
                group1.Append(textWrap1);

                picture1.Append(group1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(run1);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "008167E7", RsidRunAdditionDefault = "006B6A4E", ParagraphId = "17AFF632", TextId = "466F0C22" };

                Run run4 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run4.Append(break1);

                paragraph5.Append(run4);

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
