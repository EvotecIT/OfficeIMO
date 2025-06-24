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

        private SdtBlock CoverPageGrid {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -677494102 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "015BB951", TextId = "23C586C4" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "111F3424" };

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 16", Style = "position:absolute;margin-left:0;margin-top:0;width:422.3pt;height:760.1pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:690;mso-height-percent:960;mso-left-percent:20;mso-top-percent:20;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:690;mso-height-percent:960;mso-left-percent:20;mso-top-percent:20;mso-width-relative:page;mso-height-relative:page;v-text-anchor:middle", OptionalString = "_x0000_s1026", FillColor = "#4472c4 [3204]", Stroked = false };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nSwUGAAAAAAQABADzAAAASAUAAAAA\n"));

                V.TextBox textBox1 = new V.TextBox() { Inset = "21.6pt,1in,21.6pt" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                Caps caps1 = new Caps();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "80" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "80" };

                runProperties2.Append(caps1);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                SdtId sdtId2 = new SdtId() { Val = -1275550102 };
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "6383FA09", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Caps caps2 = new Caps();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "80" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "80" };

                paragraphMarkRunProperties1.Append(caps2);
                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Caps caps3 = new Caps();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "80" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "80" };

                runProperties3.Append(caps3);
                runProperties3.Append(color3);
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

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "591C17EA", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240" };
                Indentation indentation1 = new Indentation() { Left = "720" };
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties2.Append(color4);

                paragraphProperties2.Append(spacingBetweenLines1);
                paragraphProperties2.Append(indentation1);
                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                paragraph3.Append(paragraphProperties2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize4 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "21" };

                runProperties4.Append(color5);
                runProperties4.Append(fontSize4);
                runProperties4.Append(fontSizeComplexScript4);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Abstract" };
                SdtId sdtId3 = new SdtId() { Val = -1812170092 };
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\'", XPath = "/ns0:CoverPageProperties[1]/ns0:Abstract[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties4);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "25CE5463", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "240" };
                Indentation indentation2 = new Indentation() { Left = "1008" };
                Justification justification3 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties3.Append(color6);

                paragraphProperties3.Append(spacingBetweenLines2);
                paragraphProperties3.Append(indentation2);
                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Color color7 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize5 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "21" };

                runProperties5.Append(color7);
                runProperties5.Append(fontSize5);
                runProperties5.Append(fontSizeComplexScript5);
                Text text2 = new Text();
                text2.Text = "[Draw your reader in with an engaging abstract. It is typically a short summary of the document. When you’re ready to add your content, just click here and start typing.]";

                run3.Append(runProperties5);
                run3.Append(text2);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(run3);

                sdtContentBlock3.Append(paragraph4);

                sdtBlock3.Append(sdtProperties3);
                sdtBlock3.Append(sdtEndCharProperties3);
                sdtBlock3.Append(sdtContentBlock3);

                textBoxContent1.Append(sdtBlock2);
                textBoxContent1.Append(paragraph3);
                textBoxContent1.Append(sdtBlock3);

                textBox1.Append(textBoxContent1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle1.Append(textBox1);
                rectangle1.Append(textWrap1);

                picture1.Append(rectangle1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                Run run4 = new Run();

                RunProperties runProperties6 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties6.Append(noProof2);

                Picture picture2 = new Picture() { AnchorId = "50413575" };

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 472", Style = "position:absolute;margin-left:0;margin-top:0;width:148.1pt;height:760.3pt;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-width-percent:242;mso-height-percent:960;mso-left-percent:730;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:page;mso-width-percent:242;mso-height-percent:960;mso-left-percent:730;mso-width-relative:page;mso-height-relative:page;v-text-anchor:middle", OptionalString = "_x0000_s1027", FillColor = "#44546a [3215]", Stroked = false, StrokeWeight = "1pt" };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAlCzf6lQIAAIwFAAAOAAAAZHJzL2Uyb0RvYy54bWysVMFu2zAMvQ/YPwi6r7aDpc2MOkXQosOA\noC3WDj0rslQbk0VNUmJnXz9Kst2uK3YY5oNgiuQj+UTy/GLoFDkI61rQFS1OckqE5lC3+qmi3x6u\nP6wocZ7pminQoqJH4ejF+v27896UYgENqFpYgiDalb2paOO9KbPM8UZ0zJ2AERqVEmzHPIr2Kast\n6xG9U9kiz0+zHmxtLHDhHN5eJSVdR3wpBfe3Ujrhiaoo5ubjaeO5C2e2Pmflk2WmafmYBvuHLDrW\nagw6Q10xz8jetn9AdS234ED6Ew5dBlK2XMQasJoif1XNfcOMiLUgOc7MNLn/B8tvDvfmzobUndkC\n/+6Qkaw3rpw1QXCjzSBtF2wxcTJEFo8zi2LwhONlsVrlqzMkm6Pu0+lyuSoizxkrJ3djnf8soCPh\np6IWnymyxw5b50MCrJxMYmag2vq6VSoKoTXEpbLkwPBR/bAIj4ge7qWV0sFWQ/BK6nATC0u1xKr8\nUYlgp/RXIUlbY/aLmEjsv+cgjHOhfZFUDatFir3M8ZuiT2nFXCJgQJYYf8YeASbLBDJhpyxH++Aq\nYvvOzvnfEkvOs0eMDNrPzl2rwb4FoLCqMXKyn0hK1ASW/LAbkBt82GAZbnZQH+8ssZDGyRl+3eJD\nbpnzd8zi/ODj407wt3hIBX1FYfyjpAH78637YI9tjVpKepzHirofe2YFJeqLxoYvVgvsK5zgKH1c\nni1QsL+pdi9Vet9dAvZHgfvH8PgbHLyafqWF7hGXxybERRXTHKNXlHs7CZc+bQpcP1xsNtEMx9Yw\nv9X3hgfwwHRo1YfhkVkz9rPHUbiBaXpZ+aqtk23w1LDZe5Bt7PlnZsc3wJGPzTSup7BTXsrR6nmJ\nrn8BAAD//wMAUEsDBBQABgAIAAAAIQAjwkrH2gAAAAYBAAAPAAAAZHJzL2Rvd25yZXYueG1sTI/B\nTsMwEETvSPyDtUjcqE2kBprGqRCCW0EQ+ADH3jhR43Vku234ewwXuIy0mtHM23q3uImdMMTRk4Tb\nlQCGpL0ZyUr4/Hi+uQcWkyKjJk8o4Qsj7JrLi1pVxp/pHU9tsiyXUKyUhCGlueI86gGdiis/I2Wv\n98GplM9guQnqnMvdxAshSu7USHlhUDM+DqgP7dFJeLrbH2wQ3djrven1mtvX9uVNyuur5WELLOGS\n/sLwg5/RoclMnT+SiWySkB9Jv5q9YlMWwLocWheiBN7U/D9+8w0AAP//AwBQSwECLQAUAAYACAAA\nACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQIt\nABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAAAC8BAABfcmVscy8ucmVsc1BLAQIt\nABQABgAIAAAAIQAlCzf6lQIAAIwFAAAOAAAAAAAAAAAAAAAAAC4CAABkcnMvZTJvRG9jLnhtbFBL\nAQItABQABgAIAAAAIQAjwkrH2gAAAAYBAAAPAAAAAAAAAAAAAAAAAO8EAABkcnMvZG93bnJldi54\nbWxQSwUGAAAAAAQABADzAAAA9gUAAAAA\n"));

                V.TextBox textBox2 = new V.TextBox() { Inset = "14.4pt,,14.4pt" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Color color8 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties7.Append(color8);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Subtitle" };
                SdtId sdtId4 = new SdtId() { Val = -505288762 };
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties7);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "399DD80B", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color9 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                paragraphMarkRunProperties4.Append(color9);

                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run5 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Color color10 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties8.Append(color10);
                Text text3 = new Text();
                text3.Text = "[Document subtitle]";

                run5.Append(runProperties8);
                run5.Append(text3);

                paragraph5.Append(paragraphProperties4);
                paragraph5.Append(run5);

                sdtContentBlock4.Append(paragraph5);

                sdtBlock4.Append(sdtProperties4);
                sdtBlock4.Append(sdtEndCharProperties4);
                sdtBlock4.Append(sdtContentBlock4);

                textBoxContent2.Append(sdtBlock4);

                textBox2.Append(textBoxContent2);
                Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle2.Append(textBox2);
                rectangle2.Append(textWrap2);

                picture2.Append(rectangle2);

                run4.Append(runProperties6);
                run4.Append(picture2);

                paragraph1.Append(run1);
                paragraph1.Append(run4);
                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "36E118CA", TextId = "77777777" };

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00D6736A", RsidRunAdditionDefault = "00DE04D5", ParagraphId = "1C16249A", TextId = "28515493" };

                Run run6 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run6.Append(break1);

                paragraph7.Append(run6);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph6);
                sdtContentBlock1.Append(paragraph7);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;


            }

        }
    }
}
