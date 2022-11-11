using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithBookmarks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Bookmarks.Count == 0);
                Assert.True(document.Paragraphs.Count == 0);

                var bookmark = document.AddParagraph("Test 1").AddBookmark("Start");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 2").AddBookmark("Middle1");

                document.AddPageBreak();
                document.AddPageBreak();

                var bookmark1 = document.AddParagraph("Test 3").AddBookmark("Middle0");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 4").AddBookmark("EndOfDocument");

                Assert.True(document.Bookmarks.Count == 4);
                Assert.True(document.Paragraphs.Count == 14);

                Assert.True(document.Bookmarks[0].Name == "Start");
                Assert.True(document.Bookmarks[1].Name == "Middle1");
                Assert.True(document.Bookmarks[2].Name == "Middle0");
                Assert.True(document.Bookmarks[3].Name == "EndOfDocument");

                Assert.True(bookmark.Bookmark.Name == "Start");

                Assert.True(bookmark1.Bookmark.Name == "Middle0");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Bookmarks.Count == 4);
                Assert.True(document.Paragraphs.Count == 14);

                Assert.True(document.Bookmarks[0].Name == "Start");
                Assert.True(document.Bookmarks[1].Name == "Middle1");
                Assert.True(document.Bookmarks[2].Name == "Middle0");
                Assert.True(document.Bookmarks[3].Name == "EndOfDocument");

                //document.Bookmarks[3].Remove();


                document.AddParagraph("Test 5").AddBookmark("EndofEnds");

                Assert.True(document.Bookmarks[0].Name == "Start");
                Assert.True(document.Bookmarks[1].Name == "Middle1");
                Assert.True(document.Bookmarks[2].Name == "Middle0");
                Assert.True(document.Bookmarks[3].Name == "EndOfDocument");
                Assert.True(document.Bookmarks[4].Name == "EndofEnds");

                Assert.True(document.ParagraphsBookmarks[0].Bookmark.Name == "Start");
                Assert.True(document.ParagraphsBookmarks[1].Bookmark.Name == "Middle1");
                Assert.True(document.ParagraphsBookmarks[2].Bookmark.Name == "Middle0");
                Assert.True(document.ParagraphsBookmarks[3].Bookmark.Name == "EndOfDocument");
                Assert.True(document.ParagraphsBookmarks[4].Bookmark.Name == "EndofEnds");

                Assert.True(document.Bookmarks.Count == 5);
                Assert.True(document.Paragraphs.Count == 16);

                document.ParagraphsBookmarks[2].Bookmark.Name = "Middle6";
                document.Bookmarks[1].Name = "MiddleDocument";

                document.Save(false);
            }


            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Bookmarks.Count == 5);
                Assert.True(document.Paragraphs.Count == 16);
                Assert.True(document.ParagraphsBookmarks[0].Bookmark.Name == "Start");
                Assert.True(document.ParagraphsBookmarks[1].Bookmark.Name == "MiddleDocument");
                Assert.True(document.ParagraphsBookmarks[2].Bookmark.Name == "Middle6");
                Assert.True(document.ParagraphsBookmarks[3].Bookmark.Name == "EndOfDocument");
                Assert.True(document.ParagraphsBookmarks[4].Bookmark.Name == "EndofEnds");

                document.AddBookmark("Add bookmark straight to document. This shouldn't throw");

                Assert.True(document.Bookmarks.Count == 6);
                Assert.True(document.Paragraphs.Count == 17);
                document.Save(false);
            }
        }
    }
}
