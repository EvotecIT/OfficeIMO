using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreateAndRetrieveNumberingDefinition() {
            string filePath = Path.Combine(_directoryWithFiles, "NumberingDefinition.docx");
            int abstractId;

            using (var document = WordDocument.Create(filePath)) {
                var numbering = document.CreateNumberingDefinition();
                numbering.AddLevel(new WordListLevel(WordListLevelKind.Decimal));
                abstractId = numbering.AbstractNumberId;
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var numbering = document.GetNumberingDefinition(abstractId);
                Assert.NotNull(numbering);
                Assert.Single(numbering.Levels);
                Assert.Equal(abstractId, numbering.AbstractNumberId);
            }
        }
    }
}

