using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListStyleDetection() {
            var filePath = Path.Combine(_directoryWithFiles, "ListStyleDetection.docx");
            var styles = new[] {
                WordListStyle.Bulleted,
                WordListStyle.Numbered,
                WordListStyle.ArticleSections
            };

            using (var document = WordDocument.Create(filePath)) {
                foreach (var style in styles) {
                    var list = document.AddList(style);
                    list.AddItem("Item");
                }
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(styles.Length, document.Lists.Count);
                var detected = document.Lists.Select(l => l.Style).ToList();
                foreach (var style in styles) {
                    Assert.Contains(style, detected);
                }
            }
        }
    }
}

