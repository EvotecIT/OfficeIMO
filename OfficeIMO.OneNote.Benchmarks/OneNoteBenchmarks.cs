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
        _section = OneNoteBenchmarkCorpus.CreateSection(PageCount);
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

}
