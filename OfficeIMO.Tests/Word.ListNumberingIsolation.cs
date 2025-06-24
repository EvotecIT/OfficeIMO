using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ListNumberingIsolationBetweenLists() {
            var filePath = Path.Combine(_directoryWithFiles, "ListNumberingIsolation.docx");
            int list1Id;
            int list2Id;
            using (var document = WordDocument.Create(filePath)) {
                var list1 = document.AddList(WordListStyle.Bulleted);
                list1.AddItem("One");

                var list2 = document.AddList(WordListStyle.Bulleted);
                list2.AddItem("Two");

                list1.Bold = true;
                list1.FontSize = 20;
                list1.Color = Color.BlueViolet;

                list1Id = list1._numberId;
                list2Id = list2._numberId;

                document.Save(false);
            }

            using (var wordDoc = WordprocessingDocument.Open(filePath, false)) {
                var numbering = wordDoc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
                var inst1 = numbering.Elements<NumberingInstance>().First(n => n.NumberID!.Value == list1Id);
                var abs1 = numbering.Elements<AbstractNum>().First(a => a.AbstractNumberId!.Value == inst1.GetFirstChild<AbstractNumId>()!.Val!.Value);
                var props1 = abs1.Elements<Level>().First().NumberingSymbolRunProperties!;
                Assert.NotNull(props1.GetFirstChild<Bold>());
                Assert.NotNull(props1.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>());
                Assert.NotNull(props1.GetFirstChild<FontSize>());

                var inst2 = numbering.Elements<NumberingInstance>().First(n => n.NumberID!.Value == list2Id);
                var abs2 = numbering.Elements<AbstractNum>().First(a => a.AbstractNumberId!.Value == inst2.GetFirstChild<AbstractNumId>()!.Val!.Value);
                var props2 = abs2.Elements<Level>().First().NumberingSymbolRunProperties!;
                Assert.Null(props2.GetFirstChild<Bold>());
                Assert.Null(props2.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>());
                Assert.Null(props2.GetFirstChild<FontSize>());
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.True(document.Lists[0].Bold);
                Assert.Equal(20, document.Lists[0].FontSize);
                Assert.Equal(Color.BlueViolet, document.Lists[0].Color);

                Assert.False(document.Lists[1].Bold);
                Assert.Null(document.Lists[1].FontSize);
                Assert.Null(document.Lists[1].Color);
            }
        }
    }
}
