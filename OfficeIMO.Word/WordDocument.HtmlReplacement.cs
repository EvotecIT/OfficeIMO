using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Searches for text in the document and replaces each occurrence with an
        /// embedded HTML fragment using AltChunk.
        /// </summary>
        /// <param name="textToFind">Text to search for.</param>
        /// <param name="htmlContent">HTML fragment to insert.</param>
        /// <param name="type">Optional format type of the fragment.</param>
        /// <param name="stringComparison">String comparison option for the search.</param>
        /// <returns>The number of replacements performed.</returns>
        public int ReplaceTextWithHtmlFragment(string textToFind, string htmlContent,
            WordAlternativeFormatImportPartType type = WordAlternativeFormatImportPartType.Html,
            StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrEmpty(textToFind)) {
                throw new ArgumentNullException(nameof(textToFind));
            }

            var paragraphs = this.Paragraphs;
            var segments = SearchText(paragraphs, textToFind,
                new WordPositionInParagraph { Paragraph = 0 }, stringComparison);

            if (segments == null || segments.Count == 0) {
                return 0;
            }

            segments = segments.OrderByDescending(s => s.BeginIndex).ToList();

            foreach (var seg in segments) {
                InsertHtmlFragmentAfter(paragraphs[seg.BeginIndex], htmlContent, type);
                RemoveTextSegment(paragraphs, seg);
            }

            _document.MainDocumentPart.Document.Save();

            return segments.Count;
        }

        /// <summary>
        /// Inserts an HTML fragment after the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph after which the fragment should be inserted.</param>
        /// <param name="htmlContent">HTML content to insert.</param>
        /// <param name="type">Optional format type of the fragment.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedFragmentAfter(WordParagraph paragraph,
            string htmlContent, WordAlternativeFormatImportPartType type = WordAlternativeFormatImportPartType.Html) {
            if (paragraph == null) {
                throw new ArgumentNullException(nameof(paragraph));
            }

            return InsertHtmlFragmentAfter(paragraph, htmlContent, type);
        }

        /// <summary>
        /// Inserts an AltChunk containing HTML after the provided paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to insert after.</param>
        /// <param name="htmlContent">HTML fragment to embed.</param>
        /// <param name="type">Format type of the fragment.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        private WordEmbeddedDocument InsertHtmlFragmentAfter(WordParagraph paragraph,
            string htmlContent, WordAlternativeFormatImportPartType type) {
            MainDocumentPart mainDocPart = _document.MainDocumentPart;

            PartTypeInfo partTypeInfo = type switch {
                WordAlternativeFormatImportPartType.Rtf => AlternativeFormatImportPartType.Rtf,
                WordAlternativeFormatImportPartType.Html => AlternativeFormatImportPartType.Html,
                WordAlternativeFormatImportPartType.TextPlain => AlternativeFormatImportPartType.TextPlain,
                _ => throw new InvalidOperationException("Unsupported format type")
            };

            AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(partTypeInfo);
            string altChunkId = mainDocPart.GetIdOfPart(chunk);
            AltChunk altChunk = new AltChunk { Id = altChunkId };

            using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(htmlContent))) {
                chunk.FeedData(ms);
            }

            paragraph._paragraph.InsertAfterSelf(altChunk);

            return new WordEmbeddedDocument(this, altChunk);
        }

        /// <summary>
        /// Removes text from the paragraphs as specified by the text segment.
        /// </summary>
        /// <param name="paragraphs">Paragraph list to operate on.</param>
        /// <param name="ts">Segment describing the text range to remove.</param>
        private static void RemoveTextSegment(List<WordParagraph> paragraphs, WordTextSegment ts) {
            if (ts.BeginIndex == ts.EndIndex) {
                var p = paragraphs[ts.BeginIndex];
                var len = ts.EndChar - ts.BeginChar + 1;
                p.Text = p.Text.Remove(ts.BeginChar, len);
            } else {
                var beginPara = paragraphs[ts.BeginIndex];
                var endPara = paragraphs[ts.EndIndex];
                beginPara.Text = beginPara.Text.Substring(0, ts.BeginChar);
                endPara.Text = endPara.Text.Substring(ts.EndChar + 1);
                for (int i = ts.EndIndex - 1; i > ts.BeginIndex; i--) {
                    paragraphs[i].Remove();
                }
            }
        }
    }
}
