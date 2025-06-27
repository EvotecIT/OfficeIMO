using System.IO;
using System.Collections.Generic;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingDropDownList() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithDropDownList.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var items = new List<string> { "One", "Two" };
                var ddl = document.AddParagraph("Choose:").AddDropDownList(items, "DDL", "DDLTag");

                Assert.Single(document.DropDownLists);
                Assert.Equal(2, ddl.Items.Count);
                Assert.Equal("DDL", ddl.Alias);
                Assert.Equal("DDLTag", ddl.Tag);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.DropDownLists);
                var list = document.GetDropDownListByAlias("DDL");
                Assert.NotNull(list);
                Assert.Equal("DDLTag", document.GetDropDownListByTag("DDLTag")?.Tag);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.DropDownLists[0].Remove();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.DropDownLists);
            }
        }
    }
}
