using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word.Markdown {
    public partial class WordToMarkdownConverter {
        private static string GetParagraphText(Paragraph paragraph) {
            StringBuilder sb = new StringBuilder();
            foreach (var run in paragraph.Elements<Run>()) {
                var text = run.GetFirstChild<Text>()?.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }
                bool bold = run.RunProperties?.Bold != null;
                bool italic = run.RunProperties?.Italic != null;
                if (bold) sb.Append("**").Append(text).Append("**");
                else if (italic) sb.Append('*').Append(text).Append('*');
                else sb.Append(text);
            }
            return sb.ToString();
        }
    }
}
