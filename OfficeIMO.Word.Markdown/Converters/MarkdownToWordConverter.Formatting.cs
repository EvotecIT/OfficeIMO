using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private static string? ResolveDefaultFontFamily(MarkdownToWordOptions options) {
            if (options == null) {
                return null;
            }

            return FontResolver.Resolve(options.FontFamily) ?? options.FontFamily;
        }

        private static void ApplyBlockParagraphFormatting(WordParagraph paragraph, int quoteDepth, Omd.ColumnAlignment alignment) {
            if (quoteDepth > 0) {
                paragraph.IndentationBefore = IndentTwipsPerLevel * quoteDepth;
            }

            ApplyAlignment(alignment, paragraph);
        }

        private static void ApplyAlignment(Omd.ColumnAlignment align, WordParagraph para) {
            switch (align) {
                case Omd.ColumnAlignment.Left: para.ParagraphAlignment = JustificationValues.Left; break;
                case Omd.ColumnAlignment.Center: para.ParagraphAlignment = JustificationValues.Center; break;
                case Omd.ColumnAlignment.Right: para.ParagraphAlignment = JustificationValues.Right; break;
            }
        }
    }
}
