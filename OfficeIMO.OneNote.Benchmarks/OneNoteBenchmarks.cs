using BenchmarkDotNet.Attributes;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.OneNote.Benchmarks;

/// <summary>Tracks native offline OneNote read, write, and projection scaling.</summary>
[MemoryDiagnoser]
public class OneNoteReadWriteBenchmarks {
    private OneNoteSection _section = null!;
    private byte[] _desktopBytes = null!;
    private MemoryStream _desktopStream = null!;

    [Params(1, 25)]
    public int PageCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _section = CreateSection(PageCount);
        _desktopBytes = OneNoteSectionWriter.Write(_section);
        _desktopStream = new MemoryStream(_desktopBytes, writable: false);
    }

    [GlobalCleanup]
    public void Cleanup() => _desktopStream.Dispose();

    [Benchmark]
    public OneNoteSection ReadDesktopSection() {
        _desktopStream.Position = 0;
        return OneNoteSectionReader.Read(_desktopStream);
    }

    [Benchmark]
    public byte[] WriteDesktopSection() => OneNoteSectionWriter.Write(_section);

    [Benchmark]
    public string ProjectMarkdown() => _section.ToMarkdown();

    private static OneNoteSection CreateSection(int pageCount) {
        var section = new OneNoteSection { Name = "Benchmark section" };
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            var page = new OneNotePage { Title = "Page " + (pageIndex + 1), Level = pageIndex % 3 };
            var outline = new OneNoteOutline();
            for (int paragraphIndex = 0; paragraphIndex < 8; paragraphIndex++) {
                var paragraph = new OneNoteParagraph();
                var run = new OneNoteTextRun { Text = "Offline OneNote benchmark paragraph " + paragraphIndex + " on page " + pageIndex + "." };
                run.Style.Bold = paragraphIndex % 3 == 0;
                run.Style.Italic = paragraphIndex % 4 == 0;
                paragraph.Runs.Add(run);
                outline.Children.Add(paragraph);
            }
            page.Outlines.Add(outline);
            section.Pages.Add(page);
        }
        return section;
    }
}
