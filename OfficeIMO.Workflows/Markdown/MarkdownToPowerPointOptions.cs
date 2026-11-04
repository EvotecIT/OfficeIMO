using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Workflows;

/// <summary>Boundaries that start a new slide in a Markdown outline.</summary>
public enum MarkdownSlideBreak {
    /// <summary>A thematic break or title-level heading starts a slide.</summary>
    ThematicBreakOrHeading,
    /// <summary>Only thematic breaks start slides.</summary>
    ThematicBreak,
    /// <summary>Only title-level headings start slides.</summary>
    Heading
}

/// <summary>Controls Markdown outline projection through the existing native Markup presentation renderer.</summary>
public sealed class MarkdownToPowerPointOptions {
    /// <summary>Rejects reported mapping or native presentation loss.</summary>
    public bool FailOnLoss { get; set; }
    /// <summary>Slide boundary policy; defaults to thematic breaks and title-level headings.</summary>
    public MarkdownSlideBreak SlideBreak { get; set; }
    /// <summary>Heading level used for slide titles, from one through six.</summary>
    public int TitleHeadingLevel { get; set; } = 1;
    /// <summary>Projects blockquotes into speaker notes instead of visible slide content.</summary>
    public bool BlockquotesAsNotes { get; set; } = true;
    /// <summary>Uses title, subtitle and author front matter on the opening slide.</summary>
    public bool RenderFrontMatter { get; set; } = true;
    /// <summary>Optional native slide dimensions; omission keeps the Markup renderer's default canvas.</summary>
    public PowerPointSlideSizePreset? SlideSize { get; set; }
    /// <summary>Base directory for relative images. File conversion supplies the source directory when omitted.</summary>
    public string? BaseDirectory { get; set; }
    /// <summary>Allows local image reads. Disabled by default.</summary>
    public bool AllowLocalImages { get; set; }
    /// <summary>Local image roots. When empty, BaseDirectory is the only allowed root.</summary>
    public ISet<string> AllowedImageDirectories { get; private set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Allows bounded HTTP(S) image retrieval through the shared loader. Disabled by default.</summary>
    public bool AllowRemoteImages { get; set; }
    /// <summary>Network and byte policy for opt-in remote images.</summary>
    public OfficeRemoteImageLoadOptions? RemoteImageLoadOptions { get; set; }
    /// <summary>Allows source data URI images.</summary>
    public bool AllowDataUriImages { get; set; } = true;
    /// <summary>Maximum decoded source data URI image size.</summary>
    public long MaximumDataUriImageBytes { get; set; } = 16L * 1024 * 1024;
    /// <summary>Passes Mermaid diagrams to the existing optional native renderer; false retains them as code.</summary>
    public bool RenderMermaidDiagrams { get; set; } = true;
    /// <summary>Markdown syntax, dialect, normalization and input limits used by source conversion routes.</summary>
    public MarkdownLoadOptions? ReaderOptions { get; set; }

    /// <summary>Captures independent mutable parser, directory and network settings.</summary>
    public MarkdownToPowerPointOptions Clone() {
        if (TitleHeadingLevel < 1 || TitleHeadingLevel > 6) throw new ArgumentOutOfRangeException(nameof(TitleHeadingLevel));
        if (!Enum.IsDefined(typeof(MarkdownSlideBreak), SlideBreak)) throw new ArgumentOutOfRangeException(nameof(SlideBreak));
        if (MaximumDataUriImageBytes < 0) throw new ArgumentOutOfRangeException(nameof(MaximumDataUriImageBytes));
        var copy = (MarkdownToPowerPointOptions)MemberwiseClone();
        copy.ReaderOptions = ReaderOptions?.Clone();
        copy.AllowedImageDirectories = new HashSet<string>(AllowedImageDirectories, StringComparer.OrdinalIgnoreCase);
        if (RemoteImageLoadOptions is { } load) {
            copy.RemoteImageLoadOptions = new OfficeRemoteImageLoadOptions { MaximumBytes = load.MaximumBytes, Timeout = load.Timeout,
                MaximumRedirects = load.MaximumRedirects, AllowPrivateNetworkAddresses = load.AllowPrivateNetworkAddresses };
            foreach (string host in load.AllowedHosts) copy.RemoteImageLoadOptions.AllowedHosts.Add(host);
        }
        return copy;
    }
}
