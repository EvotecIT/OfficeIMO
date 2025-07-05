using System;
using System.Linq;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Comments {
        internal static void Example_ThreadedComments(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating threaded comments");
            string filePath = System.IO.Path.Combine(folderPath, "Threaded Comments.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Main paragraph with comment");
                paragraph.AddComment("Author1", "A1", "Top level comment");
                var parent = document.Comments.Last();
                parent.AddReply("Author2", "A2", "First reply");
                var reply = parent.AddReply("Author3", "A3", "Second reply");
                reply.AddReply("Author2", "A2", "Reply to second");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (var cmt in document.Comments) {
                    Console.WriteLine($"Comment {cmt.ParaId} text: {cmt.Text}");
                    if (cmt.ParentComment != null) {
                        Console.WriteLine($"  Parent: {cmt.ParentComment.ParaId}");
                    }
                    foreach (var rep in cmt.Replies) {
                        Console.WriteLine($"  Reply: {rep.ParaId} -> {rep.Text}");
                    }
                }
                document.Save(openWord);
            }
        }
    }
}
