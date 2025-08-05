using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word.Html {
    public partial class WordToHtmlConverter {
        private static void AppendRuns(StringBuilder sb, Paragraph paragraph, WordToHtmlOptions options, MainDocumentPart mainPart) {
            foreach (Run run in paragraph.Elements<Run>()) {
                Drawing? drawing = run.GetFirstChild<Drawing>();
                if (drawing != null) {
                    A.Blip? blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                    string? embed = blip?.Embed;
                    if (embed != null) {
                        ImagePart part = (ImagePart)mainPart.GetPartById(embed);
                        using Stream imgStream = part.GetStream();
                        using MemoryStream ms = new MemoryStream();
                        imgStream.CopyTo(ms);
                        string base64 = System.Convert.ToBase64String(ms.ToArray());
                        sb.Append($"<img src=\"data:{part.ContentType};base64,{base64}\" />");
                    }
                    continue;
                }

                string text = run.InnerText;
                string encoded = System.Net.WebUtility.HtmlEncode(text);
                RunProperties? runProps = run.RunProperties;
                string result = encoded;

                if (options.IncludeFontStyles && runProps?.RunFonts?.Ascii != null) {
                    result = $"<span style=\"font-family:{runProps.RunFonts.Ascii}\">{result}</span>";
                }

                if (runProps?.Underline != null && runProps.Underline.Val != UnderlineValues.None) {
                    result = $"<u>{result}</u>";
                }
                if (runProps?.Italic != null) {
                    result = $"<i>{result}</i>";
                }
                if (runProps?.Bold != null) {
                    result = $"<b>{result}</b>";
                }

                sb.Append(result);
            }
        }
    }
}
