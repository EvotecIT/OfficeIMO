using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static void AddBookmarkIfPresent(IElement element, WordParagraph paragraph) {
            var id = element.GetAttribute("id");
            if (string.IsNullOrEmpty(id)) {
                id = element.GetAttribute("name");
            }
            if (string.IsNullOrEmpty(id)) {
                id = element.GetAttribute("data-bookmark");
            }
            if (!string.IsNullOrEmpty(id)) {
                WordBookmark.AddBookmark(paragraph, id!);
            }
        }

        private WordParagraph AddNoteReference(WordParagraph paragraph, string text, HtmlToWordOptions options, NoteReferenceType? noteType = null) {
            var resolvedNoteType = noteType ?? options.NoteReferenceType;
            return resolvedNoteType == NoteReferenceType.Endnote
                ? paragraph.AddEndNote(text)
                : paragraph.AddFootNote(text);
        }

        private WordParagraph AddNoteReference(WordParagraph paragraph, IReadOnlyList<string> paragraphs, HtmlToWordOptions options, NoteReferenceType? noteType = null) {
            var noteReference = AddNoteReference(paragraph, paragraphs.Count == 0 ? string.Empty : paragraphs[0], options, noteType);
            if (paragraphs.Count <= 1) {
                return noteReference;
            }

            var resolvedNoteType = noteType ?? options.NoteReferenceType;
            var noteParagraphs = resolvedNoteType == NoteReferenceType.Endnote
                ? noteReference.EndNote?.Paragraphs
                : noteReference.FootNote?.Paragraphs;
            var current = noteParagraphs?.LastOrDefault();
            if (current == null) {
                return noteReference;
            }

            for (int i = 1; i < paragraphs.Count; i++) {
                current = current.AddParagraph(paragraphs[i] ?? string.Empty);
            }

            return noteReference;
        }

        private void TryLinkNoteReference(WordParagraph noteReference, string text, HtmlToWordOptions options, NoteReferenceType? noteType = null) {
            if (!options.LinkNoteUrls) {
                return;
            }
            if (!TryCreateUriFromText(text, out var uri, out var displayText)) {
                return;
            }

            var resolvedNoteType = noteType ?? options.NoteReferenceType;
            var noteParagraph = resolvedNoteType == NoteReferenceType.Endnote
                ? noteReference.EndNote?.Paragraphs?.FirstOrDefault()
                : noteReference.FootNote?.Paragraphs?.FirstOrDefault();
            if (noteParagraph == null) {
                return;
            }

            ReplaceNoteParagraphWithHyperlink(noteParagraph, uri, displayText);
        }

        private static void ReplaceNoteParagraphWithHyperlink(WordParagraph paragraph, Uri uri, string displayText) {
            if (paragraph == null) {
                return;
            }
            var runs = paragraph._paragraph.Elements<Run>().ToList();
            foreach (var run in runs) {
                if (run.GetFirstChild<FootnoteReferenceMark>() != null) {
                    continue;
                }
                if (run.GetFirstChild<EndnoteReferenceMark>() != null) {
                    continue;
                }
                run.Remove();
            }

            WordHyperLink.AddHyperLink(paragraph, displayText, uri);
        }

        private void InsertTopBookmarkIfNeeded(WordDocument doc) {
            if (!_pendingTopBookmark) {
                return;
            }

            if (doc.Bookmarks.Any(b => string.Equals(b.Name, "_top", StringComparison.OrdinalIgnoreCase))) {
                _pendingTopBookmark = false;
                return;
            }

            var firstParagraph = doc.Paragraphs.FirstOrDefault();
            if (firstParagraph == null) {
                return;
            }

            WordBookmark.AddBookmark(firstParagraph, "_top");
            _pendingTopBookmark = false;
        }

        private static bool IsInvalidHref(string href, HtmlToWordOptions options) {
            if (string.IsNullOrWhiteSpace(href)) {
                return true;
            }
            var trimmed = href.Trim();
            return !HtmlUrlPolicyEvaluator.IsAllowed(trimmed, options.HyperlinkUrlPolicy, allowEmptyFragment: false);
        }

        private static string NormalizeHref(string href) {
            var trimmed = href.Trim();
            if (trimmed.StartsWith("://", StringComparison.Ordinal)) {
                return "http" + trimmed;
            }
            if (trimmed.StartsWith("www.", StringComparison.OrdinalIgnoreCase)) {
                return "http://" + trimmed;
            }
            return trimmed;
        }

        private static bool TryCreateUriFromText(string text, out Uri uri, out string displayText) {
            uri = null!;
            displayText = text;
            if (string.IsNullOrWhiteSpace(text)) {
                return false;
            }

            var trimmed = text.Trim();
            displayText = trimmed;

            if (trimmed.StartsWith(@"\\", StringComparison.Ordinal)) {
                var normalized = "file://" + trimmed.TrimStart('\\').Replace('\\', '/');
                if (Uri.TryCreate(normalized, UriKind.Absolute, out var parsed) && parsed != null) {
                    uri = parsed;
                    return true;
                }
            }

            var candidate = NormalizeHref(trimmed);
            if (Uri.TryCreate(candidate, UriKind.Absolute, out var absolute) && absolute != null) {
                uri = absolute;
                return true;
            }

            if (Uri.TryCreate(candidate, UriKind.RelativeOrAbsolute, out var maybe) && maybe != null && maybe.IsAbsoluteUri) {
                uri = maybe;
                return true;
            }

            return false;
        }

        private static bool? GetBidiFromDir(IElement element) {
            for (IElement? current = element; current != null; current = current.ParentElement) {
                var dir = current.GetAttribute("dir");
                if (!string.IsNullOrWhiteSpace(dir)) {
                    if (string.Equals(dir, "rtl", StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                    if (string.Equals(dir, "ltr", StringComparison.OrdinalIgnoreCase)) {
                        return false;
                    }
                }
                var style = current.GetAttribute("style");
                if (TryGetDirectionFromStyle(style, out var bidi)) {
                    return bidi;
                }
            }
            return null;
        }

        private static bool TryGetDirectionFromStyle(string? style, out bool bidi) {
            bidi = false;
            if (string.IsNullOrWhiteSpace(style)) {
                return false;
            }
            foreach (var part in (style ?? string.Empty).Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length != 2) {
                    continue;
                }
                if (!string.Equals(pieces[0].Trim(), "direction", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                var value = pieces[1].Trim();
                if (string.Equals(value, "rtl", StringComparison.OrdinalIgnoreCase)) {
                    bidi = true;
                    return true;
                }
                if (string.Equals(value, "ltr", StringComparison.OrdinalIgnoreCase)) {
                    bidi = false;
                    return true;
                }
                return false;
            }
            return false;
        }

        private static void ApplyBidiIfPresent(IElement element, WordParagraph paragraph) {
            var bidi = GetBidiFromDir(element);
            if (bidi.HasValue) {
                paragraph.BiDi = bidi.Value;
            }
        }
    }
}
