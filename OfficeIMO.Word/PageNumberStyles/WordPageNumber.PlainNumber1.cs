using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Represents a page-numbering building block.
/// </summary>
public partial class WordPageNumber {
    private static SdtBlock PlainNumber1 {
        get {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtId sdtId1 = new SdtId();

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            sdtEndCharProperties1.Append(runProperties1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "000003E1", RsidRunAdditionDefault = "000003E1", ParagraphId = "2D5F018E", TextId = "0961F169" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run1.Append(fieldChar1);

            Run run2 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

            run2.Append(fieldCode1);

            Run run3 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run3.Append(fieldChar2);

            Run run4 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            Text text1 = new Text();
            text1.Text = "2";

            run4.Append(runProperties2);
            run4.Append(text1);

            Run run5 = new Run();

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run5.Append(runProperties3);
            run5.Append(fieldChar3);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtEndCharProperties1);
            sdtBlock1.Append(sdtContentBlock1);
            return sdtBlock1;


        }
    }
}
