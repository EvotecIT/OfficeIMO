using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

/// <summary>HTML-backed EPUB chapter image-export options.</summary>
public sealed class EpubImageExportOptions : HtmlRenderOptions {
    /// <summary>Creates one continuous preview per chapter; paged output remains available through the shared mode setting.</summary>
    public EpubImageExportOptions() => Mode = HtmlRenderMode.Continuous;

    /// <summary>Zero-based first chapter to export.</summary>
    public int ChapterIndex { get; set; }

    /// <summary>Maximum chapter count, or null for every remaining chapter.</summary>
    public int? ChapterCount { get; set; }

    /// <summary>Adds a chapter heading when the raw chapter markup does not already expose one.</summary>
    public bool IncludeChapterTitle { get; set; } = true;

    /// <summary>Maps EPUB package diagnostics into every exported result.</summary>
    public bool IncludePackageDiagnostics { get; set; } = true;

    /// <summary>Creates an independent EPUB options snapshot.</summary>
    public EpubImageExportOptions CloneEpub() {
        EpubImageExportOptions clone = CopyTo(new EpubImageExportOptions());
        clone.ChapterIndex = ChapterIndex;
        clone.ChapterCount = ChapterCount;
        clone.IncludeChapterTitle = IncludeChapterTitle;
        clone.IncludePackageDiagnostics = IncludePackageDiagnostics;
        return clone;
    }

    /// <inheritdoc />
    public override HtmlRenderOptions Clone() => CloneEpub();
}
