using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.OneNote.Benchmarks;

internal static class OneNoteBenchmarkCorpus {
    internal const string EditMarker = "OfficeIMO benchmark edit marker";
    internal static readonly OneNoteBenchmarkScale[] Scales = {
        new("Small", 1),
        new("Normal", 25),
        new("Large", 100)
    };

    internal static OneNoteBenchmarkScale Get(string name) =>
        Scales.FirstOrDefault(scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown OneNote benchmark scale: " + name, nameof(name));

    internal static OneNoteSection CreateSection(int pageCount) {
        var section = new OneNoteSection { Name = "Benchmark section" };
        for (var pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            section.Pages.Add(CreatePage(pageIndex));
        }
        return section;
    }

    internal static OneNotePage CreateEditPage(int pageIndex) {
        OneNotePage page = CreatePage(pageIndex);
        page.Title = EditMarker;
        ((OneNoteParagraph)page.Outlines[0].Children[0]).Runs[0].Text = EditMarker;
        return page;
    }

    internal static OneNoteCorpusObservation Observe(OneNoteSection section) {
        var paragraphCount = 0;
        bool containsEditMarker = false;
        var canonical = new StringBuilder();
        Append(canonical, section.Name);
        foreach (OneNotePage page in section.Pages) {
            Append(canonical, page.Title);
            Append(canonical, page.Level);
            Append(canonical, page.CreatedUtc?.ToUniversalTime().Ticks ?? -1L);
            Append(canonical, page.LastModifiedUtc?.ToUniversalTime().Ticks ?? -1L);
            Append(canonical, page.Outlines.Count);
            containsEditMarker |= string.Equals(page.Title, EditMarker, StringComparison.Ordinal);
            foreach (OneNoteOutline outline in page.Outlines) {
                Append(canonical, outline.Children.Count);
                foreach (OneNoteElement element in outline.Children) {
                    Append(canonical, (int)element.Kind);
                    if (element is not OneNoteParagraph paragraph) {
                        throw new InvalidOperationException("The OneNote benchmark corpus contains an unexpected non-paragraph element.");
                    }
                    paragraphCount++;
                    Append(canonical, paragraph.Runs.Count);
                    foreach (OneNoteTextRun run in paragraph.Runs) {
                        Append(canonical, run.Text);
                        Append(canonical, run.Style.Bold);
                        Append(canonical, run.Style.Italic);
                        containsEditMarker |= string.Equals(run.Text, EditMarker, StringComparison.Ordinal);
                    }
                }
            }
        }
        using SHA256 sha256 = SHA256.Create();
        string structuralFingerprint = Convert.ToHexString(sha256.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString())));
        return new OneNoteCorpusObservation(section.Pages.Count, paragraphCount, structuralFingerprint, containsEditMarker);
    }

    private static void Append(StringBuilder builder, string value) =>
        builder.Append(value.Length).Append(':').Append(value).Append(';');

    private static void Append(StringBuilder builder, int value) =>
        builder.Append(value).Append(';');

    private static void Append(StringBuilder builder, long value) =>
        builder.Append(value).Append(';');

    private static void Append(StringBuilder builder, bool? value) =>
        builder.Append(value.HasValue ? value.Value ? '1' : '0' : '-').Append(';');

    private static OneNotePage CreatePage(int pageIndex) {
        var page = new OneNotePage {
            Title = "Page " + (pageIndex + 1),
            Level = pageIndex % 3,
            CreatedUtc = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMinutes(pageIndex),
            LastModifiedUtc = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMinutes(pageIndex + 1)
        };
        var outline = new OneNoteOutline();
        for (var paragraphIndex = 0; paragraphIndex < 8; paragraphIndex++) {
            var paragraph = new OneNoteParagraph();
            var run = new OneNoteTextRun {
                Text = "Offline OneNote benchmark paragraph " + paragraphIndex + " on page " + pageIndex + "."
            };
            run.Style.Bold = paragraphIndex % 3 == 0;
            run.Style.Italic = paragraphIndex % 4 == 0;
            paragraph.Runs.Add(run);
            outline.Children.Add(paragraph);
        }
        page.Outlines.Add(outline);
        return page;
    }
}

internal sealed record OneNoteBenchmarkScale(string Name, int PageCount);

internal sealed record OneNoteCorpusObservation(
    int PageCount,
    int ParagraphCount,
    string StructuralFingerprint,
    bool ContainsEditMarker);
