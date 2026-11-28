using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentReadBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixturePath = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENTS_FIXTURE");
        _fixture = fixturePath == null ? CreateFixture() : File.ReadAllBytes(fixturePath);
        if (ReadComments() != 24 || ReadEmptySlides() != 0) {
            throw new InvalidOperationException("The comment fixture did not have the expected review data.");
        }
    }

    [Benchmark]
    public int ReadComments() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        int count = 0;
        for (int index = 0; index < 12; index++) {
            foreach (PowerPointClassicComment comment in presentation.GetClassicComments(presentation.Slides[index])) {
                if (comment.Index == 0 || comment.Created is not { Kind: DateTimeKind.Utc } ||
                    comment.Text.Length == 0 || comment.Author.Name.Length == 0 || comment.X == comment.Y) {
                    throw new InvalidOperationException("The comment fixture did not preserve its fields.");
                }
                count++;
            }
        }
        return count;
    }

    [Benchmark]
    public int ReadEmptySlides() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        int count = 0;
        for (int index = 12; index < presentation.Slides.Count; index++) {
            count += presentation.GetClassicComments(presentation.Slides[index]).Count;
        }
        return count;
    }

    internal static byte[] CreateFixture() {
        using var presentation = PowerPointDocument.Create();
        var editor = new PowerPointCommentAuthor("Presentation Editor", "PE");
        var reviewer = new PowerPointCommentAuthor("Presentation Reviewer", "PR");
        var date = new DateTime(2026, 4, 18, 12, 30, 0, DateTimeKind.Utc);
        for (int index = 0; index < 24; index++) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Slide " + (index + 1));
            if (index >= 12) continue;
            presentation.AddClassicComment(slide, editor, "An editorial note", OfficeLength.Points(12), OfficeLength.Points(30), date);
            presentation.AddClassicComment(slide, reviewer, "A reviewer response", OfficeLength.Points(130), OfficeLength.Points(56), date.AddHours(1));
        }
        return presentation.ToBytes();
    }
}
