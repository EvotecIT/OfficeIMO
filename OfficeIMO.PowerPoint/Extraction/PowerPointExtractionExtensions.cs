using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
public static class PowerPointExtractionExtensions {
    /// <summary>
    /// Options controlling PowerPoint extraction behavior.
    /// </summary>
    public sealed class PowerPointExtractOptions {
        /// <summary>
        /// When true, include speaker notes in output. Default: true.
        /// </summary>
        public bool IncludeNotes { get; set; } = true;
    }

    /// <summary>
    /// Extracts a presentation into slide-aligned chunks (one chunk per slide by default).
    /// </summary>
    public static IEnumerable<PowerPointExtractChunk> ExtractMarkdownChunks(
        this PowerPointPresentation presentation,
        PowerPointExtractOptions? extract = null,
        PowerPointExtractChunkingOptions? chunking = null,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        extract ??= new PowerPointExtractOptions();
        chunking ??= new PowerPointExtractChunkingOptions();
        if (chunking.MaxChars < 256) chunking.MaxChars = 256;

        for (int i = 0; i < presentation.Slides.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();

            var slide = presentation.Slides[i];
            int slideNumber = i + 1;

            var md = new StringBuilder();
            md.Append("## Slide ").AppendLine(slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
            md.AppendLine();

            // TextBoxes in shape order.
            foreach (var tb in slide.TextBoxes) {
                cancellationToken.ThrowIfCancellationRequested();
                var text = (tb.Text ?? string.Empty).Trim();
                if (text.Length == 0) continue;
                md.AppendLine(text);
                md.AppendLine();
            }

            if (extract.IncludeNotes) {
                var notes = (TryReadNotesTextNoCreate(slide) ?? string.Empty).Trim();
                if (notes.Length > 0) {
                    md.AppendLine("### Notes");
                    md.AppendLine();
                    md.AppendLine(notes);
                    md.AppendLine();
                }
            }

            var markdown = md.ToString().TrimEnd();
            if (markdown.Length > chunking.MaxChars) {
                markdown = markdown.Substring(0, chunking.MaxChars) + "\n\n<!-- truncated -->";
            }

            var id = BuildStableId("ppt-md", sourcePath, slideNumber);
            yield return new PowerPointExtractChunk {
                Id = id,
                Location = new PowerPointExtractLocation {
                    Path = sourcePath,
                    Slide = slideNumber,
                    BlockIndex = i
                },
                Text = markdown,
                Markdown = markdown
            };
        }
    }

    private static string? TryReadNotesTextNoCreate(PowerPointSlide slide) {
        // Avoid side effects: PowerPointSlide.Notes.Text will create a NotesSlidePart if absent.
        // For extraction we only read notes when they already exist.
        try {
            var notesPart = slide.SlidePart.NotesSlidePart;
            var notesSlide = notesPart?.NotesSlide;
            if (notesSlide == null) return null;

            Shape? shape = notesSlide.CommonSlideData?.ShapeTree?.GetFirstChild<Shape>();
            A.Paragraph? paragraph = shape?.TextBody?.GetFirstChild<A.Paragraph>();
            A.Run? run = paragraph?.GetFirstChild<A.Run>();
            A.Text? text = run?.GetFirstChild<A.Text>();
            return text?.Text ?? string.Empty;
        } catch {
            return null;
        }
    }

    private static string BuildStableId(string kind, string? path, int slideNumber) {
        var safe = string.IsNullOrWhiteSpace(path) ? "memory" : System.IO.Path.GetFileName(path!.Trim());
        return $"{kind}:{safe}:s{slideNumber}";
    }
}

