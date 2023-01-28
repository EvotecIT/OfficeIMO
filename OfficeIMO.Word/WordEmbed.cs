using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace OfficeIMO.Word {
    public class WordEmbed {

        private string GetAltChunkId(WordDocument wordDoc) {
            int id = 1;
            string altChunkId = "AltChunkId" + id;

            try {
                while (wordDoc._document.MainDocumentPart.GetPartById(altChunkId) != null) {
                    id++;
                    altChunkId = "AltChunkId" + id;
                }
            } catch {

            }
            return altChunkId;


        }

        public WordEmbed(WordDocument wordDocument, string fileName, string description) {

            FileInfo fileInfo = new FileInfo(fileName);
            AlternativeFormatImportPartType partType;
            if (fileInfo.Extension == ".rtf") {
                partType = AlternativeFormatImportPartType.Rtf;
            } else if (fileInfo.Extension == ".html") {
                partType = AlternativeFormatImportPartType.Html;
            } else {
                throw new Exception("Only RTF files are supported for now :-)");
            }


            AltChunk altChunk = new AltChunk {
                Id = GetAltChunkId(wordDocument)
            };

            //string altChunkId = "AltChunkId5";

            MainDocumentPart mainDocPart = wordDocument._document.MainDocumentPart;

            AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(partType, altChunk.Id);

            var documentContent = File.ReadAllText(fileName, Encoding.ASCII);

            using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(documentContent))) {
                chunk.FeedData(ms);
            }


            // Embed AltChunk after the last paragraph.
            mainDocPart.Document.Body.InsertAfter(altChunk, mainDocPart.Document.Body.Elements<Paragraph>().Last());
            mainDocPart.Document.Save();

        }
    }
}
