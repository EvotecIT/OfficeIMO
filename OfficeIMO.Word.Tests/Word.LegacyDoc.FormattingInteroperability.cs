using OfficeIMO.Drawing.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesMultipleCharacterFormattingPages() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const int runCount = 140;

            try {
                using (WordDocument document = WordDocument.Create()) {
                    WordParagraph paragraph = document.AddParagraph();
                    for (int index = 0; index < runCount; index++) {
                        WordParagraph run = paragraph.AddText($"{index:D3}|");
                        if (index % 2 == 0) {
                            run.SetBold();
                        } else {
                            run.SetItalic();
                        }
                    }

                    document.Save(docPath);
                }

                Assert.True(
                    OfficeCompoundFileReader.TryRead(File.ReadAllBytes(docPath), out OfficeCompoundFile? compoundFile, out string? error),
                    error);
                byte[] wordDocumentStream = compoundFile!.Streams["WordDocument"];
                int characterFormattingPlcLength = BitConverter.ToInt32(wordDocumentStream, 0xFE);
                Assert.True(characterFormattingPlcLength > 12);
                Assert.Equal(0, (characterFormattingPlcLength - sizeof(int)) % (sizeof(int) * 2));

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.Equal(runCount, reloaded.Paragraphs.Count);
                for (int index = 0; index < runCount; index++) {
                    WordParagraph run = reloaded.Paragraphs[index];
                    Assert.Equal($"{index:D3}|", run.Text);
                    Assert.Equal(index % 2 == 0, run.Bold);
                    Assert.Equal(index % 2 != 0, run.Italic);
                }
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }
    }
}
