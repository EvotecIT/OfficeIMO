using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_CreatingWordDocumentWithEmbeddedDocuments() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments.docx");
        string htmlFilePath = System.IO.Path.Combine(_directoryDocuments, "SampleFileHTML.html");
        string rtfFilePath = System.IO.Path.Combine(_directoryDocuments, "SampleFileRTF.rtf");


        using (var document = WordDocument.Create(filePath)) {

            Assert.True(document.EmbeddedDocuments.Count == 0);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 0);

            document.AddParagraph("Add RTF document in front of the document");

            document.AddEmbeddedDocument(rtfFilePath);

            Assert.True(document.EmbeddedDocuments.Count == 1);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 1);
            Assert.True(document.EmbeddedDocuments[0].ContentType == "application/rtf");

            document.AddPageBreak();

            document.AddParagraph("Add HTML document as last in the document");

            document.AddEmbeddedDocument(htmlFilePath);

            Assert.True(document.EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments.docx"))) {
            Assert.True(document.EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.EmbeddedDocuments[0].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");

            var tempFilePath1 = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments.rtf");
            var tempFilePath2 = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments.html");

            document.EmbeddedDocuments[0].Save(tempFilePath1);
            document.EmbeddedDocuments[1].Save(tempFilePath2);

            Assert.True(File.Exists(tempFilePath1));
            Assert.True(File.Exists(tempFilePath2));

            FileInfo info1 = new FileInfo(tempFilePath1);
            FileInfo info2 = new FileInfo(tempFilePath2);
            FileInfo infoRtf = new FileInfo(rtfFilePath);
            FileInfo infoHtml = new FileInfo(htmlFilePath);

            Assert.True(info1.Length == infoRtf.Length);
            Assert.True(info2.Length == infoHtml.Length);

            document.AddEmbeddedDocument(rtfFilePath);

            Assert.True(document.EmbeddedDocuments[0].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[2].ContentType == "application/rtf");

            Assert.True(document.EmbeddedDocuments.Count == 3);

            document.AddSection();

            Assert.True(document.EmbeddedDocuments.Count == 3);

            document.AddEmbeddedDocument(rtfFilePath);

            Assert.True(document.EmbeddedDocuments.Count == 4);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 3);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 1);

            document.AddEmbeddedDocument(rtfFilePath);

            Assert.True(document.EmbeddedDocuments.Count == 5);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 3);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 2);

            Assert.True(document.EmbeddedDocuments[0].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[2].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[3].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[4].ContentType == "application/rtf");

            var tempFilePath3 = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments3.rtf");
            var tempFilePath4 = Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments4.rtf");


            document.Sections[1].EmbeddedDocuments[0].Save(tempFilePath3);
            document.Sections[1].EmbeddedDocuments[1].Save(tempFilePath4);

            Assert.True(File.Exists(tempFilePath3));
            Assert.True(File.Exists(tempFilePath4));

            FileInfo info3 = new FileInfo(tempFilePath3);
            FileInfo info4 = new FileInfo(tempFilePath4);

            Assert.True(info3.Length == infoRtf.Length);
            Assert.True(info4.Length == infoRtf.Length);

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithEmbeddedDocuments.docx"))) {
            Assert.True(document.EmbeddedDocuments.Count == 5);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 3);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 2);

            Assert.True(document.EmbeddedDocuments[0].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[2].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[3].ContentType == "application/rtf");
            Assert.True(document.EmbeddedDocuments[4].ContentType == "application/rtf");

            var list1 = document._document.MainDocumentPart.AlternativeFormatImportParts;
            Assert.True(list1.Count() == 5);

            // lets delete last 3 embedded documents 
            document.EmbeddedDocuments[2].Remove();
            // since we deleted 3rd document, 4th document is now 3rd
            document.EmbeddedDocuments[2].Remove();
            // since we deleted 3rd document, 5th document is now 3rd
            document.EmbeddedDocuments[2].Remove();

            Assert.True(document.EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 0);

            var list2 = document._document.MainDocumentPart.AlternativeFormatImportParts;
            Assert.True(list2.Count() == 2);


            var htmlContent = """
                              <html lang="en">
                              <P>This is a paragraph.</P>
                              <P>This is another paragraph.</P>
                              <P>This is a paragraph with <STRONG>bold</STRONG> text.</P>
                              <P>This is a paragraph with <EM>italic</EM> text.</P>
                              <ul>
                                <li>Item 1</li>
                                <li>Item 2</li>
                                <li>Item 3</li>
                              </ul>
                              <ol>
                                <li>Item 1</li>
                                <li>Item 2</li>
                                <li>Item 3</li>
                              </ol>
                              <P>This is a paragraph with a <A href=""https://www.google.com"">link</A>.</P>
                              </html>
                              """;

            document.AddEmbeddedFragment(htmlContent, WordAlternativeFormatImportPartType.Html);

            Assert.True(document.EmbeddedDocuments.Count == 3);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 1);

            var list3 = document._document.MainDocumentPart.AlternativeFormatImportParts;
            Assert.True(list3.Count() == 3);
        }
    }

    [Fact]
    public void Test_LoadingWordDocumentWithEmbeddedDocumentsAndBrokenSections() {
        // Those section have no RSID so comparison of sections is very risky
        string rtfFilePath = System.IO.Path.Combine(_directoryDocuments, "SampleFileRTF.rtf");

        using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmbeddedFileBrokenSections.docx"))) {
            Assert.True(document.EmbeddedDocuments.Count == 4);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 2);
            Assert.True(document.EmbeddedDocuments[0].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[2].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[3].ContentType == "text/html");

            document.AddEmbeddedDocument(rtfFilePath);

            Assert.True(document.EmbeddedDocuments[4].ContentType == "application/rtf");

            Assert.True(document.EmbeddedDocuments.Count == 5);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 3);

            document.Save();
        }

        using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "EmbeddedFileBrokenSections.docx"))) {
            Assert.True(document.EmbeddedDocuments.Count == 5);
            Assert.True(document.Sections[0].EmbeddedDocuments.Count == 2);
            Assert.True(document.Sections[1].EmbeddedDocuments.Count == 3);

            Assert.True(document.EmbeddedDocuments[0].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[1].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[2].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[3].ContentType == "text/html");
            Assert.True(document.EmbeddedDocuments[4].ContentType == "application/rtf");
        }
    }
}
