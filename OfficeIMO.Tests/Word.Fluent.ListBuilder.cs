using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentListBuilder() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentListBuilder.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .List(l => l.Numbered().StartAt(3)
                                 .Item("First")
                                 .Item("Second")
                                 .Indent().Item("Second.Child"))
                    .List(l => l.Bulleted()
                                 .Item("Alpha")
                                 .Item("Beta").Indent().Item("Beta.Child").Outdent()
                                 .Item("Gamma"))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Lists.Count);

                var numbered = document.Lists[0];
                Assert.Equal(3, numbered.Numbering.Levels[0].StartNumberingValue);
                Assert.Equal(3, numbered.ListItems.Count);
                Assert.Equal(0, numbered.ListItems[0].ListItemLevel);
                Assert.Equal(0, numbered.ListItems[1].ListItemLevel);
                Assert.Equal(1, numbered.ListItems[2].ListItemLevel);

                var bulleted = document.Lists[1];
                Assert.Equal(4, bulleted.ListItems.Count);
                Assert.Equal(0, bulleted.ListItems[0].ListItemLevel);
                Assert.Equal(0, bulleted.ListItems[1].ListItemLevel);
                Assert.Equal(1, bulleted.ListItems[2].ListItemLevel);
                Assert.Equal(0, bulleted.ListItems[3].ListItemLevel);
            }
        }
    }
}
