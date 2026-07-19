using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>Fluent batch image export for EPUB chapters.</summary>
public sealed class EpubImageExportBuilder :
    OfficeImageExportBatchBuilder<EpubImageExportBuilder, EpubImageExportOptions> {
    internal EpubImageExportBuilder(
        EpubDocument source,
        EpubImageExportOptions? options)
        : base(
            options?.CloneEpub() ?? new EpubImageExportOptions(),
            (format, effective) => source.ExportImages(
                format,
                effective),
            (format, effective, consumer, cancellationToken) =>
                source.ExportImages(
                    format,
                    consumer,
                    effective,
                    cancellationToken),
            (format, effective, consumer, cancellationToken) =>
                source.ExportImagesAsync(
                    format,
                    consumer,
                    effective,
                    cancellationToken)) {
        if (source == null) throw new ArgumentNullException(nameof(source));
    }

    /// <summary>Starts at a zero-based chapter index.</summary>
    public EpubImageExportBuilder FromChapter(int chapterIndex) {
        if (chapterIndex < 0) {
            throw new ArgumentOutOfRangeException(nameof(chapterIndex));
        }
        Options.ChapterIndex = chapterIndex;
        return this;
    }

    /// <summary>Limits the number of chapters exported.</summary>
    public EpubImageExportBuilder TakeChapters(int chapterCount) {
        if (chapterCount < 1) {
            throw new ArgumentOutOfRangeException(nameof(chapterCount));
        }
        Options.ChapterCount = chapterCount;
        return this;
    }

    /// <summary>Uses one continuous image surface per chapter.</summary>
    public EpubImageExportBuilder Continuous() {
        Options.Mode = HtmlRenderMode.Continuous;
        return this;
    }

    /// <summary>Uses physical paged layout for each chapter.</summary>
    public EpubImageExportBuilder Paged() {
        Options.Mode = HtmlRenderMode.Paged;
        return this;
    }
}
