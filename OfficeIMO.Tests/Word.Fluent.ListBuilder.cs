using DocumentFormat.OpenXml.Wordprocessing;
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
                                 .Level(1).Item("Second.Child")
                                 .Level(2).Item("Second.Child.Grandchild")
                                 .Level(0).Item("Third"))
                    .List(l => l.Bulleted()
                                 .Item("Alpha")
                                 .Item("Beta").Indent().Item("Beta.Child").Outdent()
                                 .Item("Gamma"))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Lists.Count);

                var numbered = document.Lists[0];
                Assert.Equal(3, numbered.Numbering.Levels[0].StartNumberingValue);
                Assert.Equal(5, numbered.ListItems.Count);
                Assert.Equal(0, numbered.ListItems[0].ListItemLevel);
                Assert.Equal(0, numbered.ListItems[1].ListItemLevel);
                Assert.Equal(1, numbered.ListItems[2].ListItemLevel);
                Assert.Equal(2, numbered.ListItems[3].ListItemLevel);
                Assert.Equal(0, numbered.ListItems[4].ListItemLevel);

                var bulleted = document.Lists[1];
                Assert.Equal(4, bulleted.ListItems.Count);
                Assert.Equal(0, bulleted.ListItems[0].ListItemLevel);
                Assert.Equal(0, bulleted.ListItems[1].ListItemLevel);
                Assert.Equal(1, bulleted.ListItems[2].ListItemLevel);
                Assert.Equal(0, bulleted.ListItems[3].ListItemLevel);
            }
        }

        [Fact]
        public void Test_FluentListBuilder_CustomFormats() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentListBuilderFormats.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .List(l => l.Numbered().NumberFormat(NumberFormatValues.LowerRoman)
                                     .Item("first")
                                     .Item("second"))
                    .List(l => l.Bulleted().BulletCharacter("\u2736")
                                     .Item("alpha")
                                     .Item("beta"))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Lists.Count);

                var roman = document.Lists[0];
                Assert.Equal(NumberFormatValues.LowerRoman, roman.Numbering.Levels[0]._level.NumberingFormat.Val!.Value);

                var custom = document.Lists[1];
                Assert.Equal("\u2736", custom.Numbering.Levels[0].LevelText);
            }
        }
    }
}
