using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Services.Markdown;

internal interface IMarkdownReaderOptionSource
{
    MarkdownReaderOptions? ReaderOptions { get; }
    MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; }
    string? BaseUri { get; }
    int? MaxInputCharacters { get; }
    MarkdownInputNormalizationPreset? NormalizeInput { get; }
    bool? DisallowFileUrls { get; }
    bool? AllowDataUrls { get; }
    bool? AllowMailtoUrls { get; }
    bool? AllowProtocolRelativeUrls { get; }
    bool? RestrictUrlSchemes { get; }
    string[]? AllowedUrlScheme { get; }
}

internal interface IMarkdownWriteOptionSource
{
    MarkdownWriteOptions? WriteOptions { get; }
    OfficeMarkdownWriteProfile? WriteProfile { get; }
    MarkdownImageRenderingMode? ImageRenderingMode { get; }
    string? LineEnding { get; }
    string? UnorderedListMarker { get; }
}

internal interface IMarkdownPdfOptionSource
{
    MarkdownToPdfOptions? MarkdownPdfOptions { get; }
    OfficeIMO.Pdf.PdfOptions? PdfOptions { get; }
    OfficeVisualThemeKind? PdfTheme { get; }
    string? PdfFontFamily { get; }
    string? PdfTitle { get; }
    string? PdfAuthor { get; }
    string? PdfSubject { get; }
    string? PdfKeywords { get; }
    string? PdfBaseDirectory { get; }
    bool? PdfApplyWordLikeTheme { get; }
    bool? PdfIncludeLocalImages { get; }
    bool? PdfIncludeDataUriImages { get; }
    bool? PdfRestrictLocalImagesToBaseDirectory { get; }
    int? PdfMaximumDataUriImageBytes { get; }
    double? PdfDefaultImageWidth { get; }
    double? PdfDefaultImageHeight { get; }
    MarkdownPdfFrontMatterRenderMode? PdfFrontMatterRenderMode { get; }
    bool? PdfUseFrontMatterVisualTheme { get; }
    bool? PdfUseFrontMatterMetadata { get; }
    bool? PdfUseFirstHeadingAsTitle { get; }
    bool? PdfCreateOutlineFromHeadings { get; }
    string? PdfWarningVariable { get; }
    string? PdfConversionReportVariable { get; }
}

internal static class MarkdownOptionUtilities
{
    internal static MarkdownReaderOptions? BuildReaderOptions(IMarkdownReaderOptionSource source)
    {
        if (source.ReaderOptions != null && source.Profile.HasValue)
        {
            throw new PSArgumentException("Specify either -ReaderOptions or -Profile, not both.");
        }

        if (source.ReaderOptions == null && !source.Profile.HasValue && !HasReaderOverrides(source))
        {
            return null;
        }

        var options = source.ReaderOptions
            ?? (source.Profile.HasValue
                ? MarkdownReaderOptions.CreateProfile(source.Profile.Value)
                : MarkdownReaderOptions.CreateOfficeIMOProfile());

        if (!string.IsNullOrWhiteSpace(source.BaseUri))
        {
            options.BaseUri = source.BaseUri;
        }

        if (source.MaxInputCharacters.HasValue)
        {
            options.MaxInputCharacters = source.MaxInputCharacters.Value;
        }

        if (source.NormalizeInput.HasValue)
        {
            options.InputNormalization.ApplyPreset(source.NormalizeInput.Value);
        }

        if (source.DisallowFileUrls.HasValue)
        {
            options.DisallowFileUrls = source.DisallowFileUrls.Value;
        }

        if (source.AllowDataUrls.HasValue)
        {
            options.AllowDataUrls = source.AllowDataUrls.Value;
        }

        if (source.AllowMailtoUrls.HasValue)
        {
            options.AllowMailtoUrls = source.AllowMailtoUrls.Value;
        }

        if (source.AllowProtocolRelativeUrls.HasValue)
        {
            options.AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls.Value;
        }

        if (source.RestrictUrlSchemes.HasValue)
        {
            options.RestrictUrlSchemes = source.RestrictUrlSchemes.Value;
        }

        if (source.AllowedUrlScheme is { Length: > 0 })
        {
            options.AllowedUrlSchemes = source.AllowedUrlScheme;
            options.RestrictUrlSchemes = true;
        }

        return options;
    }

    internal static MarkdownWriteOptions? BuildWriteOptions(IMarkdownWriteOptionSource source)
    {
        if (source.WriteOptions == null && !source.WriteProfile.HasValue && !source.ImageRenderingMode.HasValue
            && string.IsNullOrWhiteSpace(source.LineEnding) && string.IsNullOrWhiteSpace(source.UnorderedListMarker))
        {
            return null;
        }

        var options = source.WriteOptions?.Clone() ?? CreateWriteProfile(source.WriteProfile);

        if (source.ImageRenderingMode.HasValue)
        {
            options.ImageRenderingMode = source.ImageRenderingMode.Value;
        }

        if (!string.IsNullOrWhiteSpace(source.LineEnding))
        {
            options.OutputLineEnding = ResolveLineEnding(source.LineEnding!);
        }

        if (!string.IsNullOrWhiteSpace(source.UnorderedListMarker))
        {
            var marker = source.UnorderedListMarker!.Trim();
            if (marker.Length != 1)
            {
                throw new PSArgumentException("Unordered list marker must be '-', '*', or '+'.", nameof(source.UnorderedListMarker));
            }

            options.UnorderedListMarker = marker[0];
        }

        return options;
    }

    internal static MarkdownToPdfOptions BuildPdfOptions(IMarkdownPdfOptionSource source, PSCmdlet command, string? fallbackBaseDirectory)
    {
        var options = source.MarkdownPdfOptions?.Clone() ?? new MarkdownToPdfOptions();

        if (source.PdfOptions != null) options.PdfOptions = source.PdfOptions;
        if (source.PdfTheme.HasValue) options.Theme = MarkdownVisualTheme.Create(source.PdfTheme.Value);
        if (!string.IsNullOrWhiteSpace(source.PdfFontFamily)) options.FontFamily = source.PdfFontFamily;
        if (!string.IsNullOrWhiteSpace(source.PdfTitle)) options.Title = source.PdfTitle;
        if (!string.IsNullOrWhiteSpace(source.PdfAuthor)) options.Author = source.PdfAuthor;
        if (!string.IsNullOrWhiteSpace(source.PdfSubject)) options.Subject = source.PdfSubject;
        if (!string.IsNullOrWhiteSpace(source.PdfKeywords)) options.Keywords = source.PdfKeywords;
        if (source.PdfApplyWordLikeTheme.HasValue) options.ApplyDefaultTheme = source.PdfApplyWordLikeTheme.Value;

        if (!string.IsNullOrWhiteSpace(source.PdfBaseDirectory))
        {
            options.BaseDirectory = PdfCommandUtilities.ResolvePath(command, source.PdfBaseDirectory!);
        }
        else if (source.PdfIncludeLocalImages == true && !string.IsNullOrWhiteSpace(fallbackBaseDirectory))
        {
            options.BaseDirectory = Path.GetFullPath(fallbackBaseDirectory!);
        }

        if (source.PdfIncludeLocalImages.HasValue) options.ResourcePolicy.AllowLocalFileAccess = source.PdfIncludeLocalImages.Value;
        if (source.PdfIncludeDataUriImages.HasValue) options.ResourcePolicy.AllowDataUris = source.PdfIncludeDataUriImages.Value;
        if (source.PdfRestrictLocalImagesToBaseDirectory.HasValue) options.RestrictLocalImagesToBaseDirectory = source.PdfRestrictLocalImagesToBaseDirectory.Value;
        if (source.PdfMaximumDataUriImageBytes.HasValue) options.MaximumDataUriImageBytes = source.PdfMaximumDataUriImageBytes.Value;
        if (source.PdfDefaultImageWidth.HasValue) options.DefaultImageWidth = source.PdfDefaultImageWidth.Value;
        if (source.PdfDefaultImageHeight.HasValue) options.DefaultImageHeight = source.PdfDefaultImageHeight.Value;
        if (source.PdfFrontMatterRenderMode.HasValue) options.FrontMatterRenderMode = source.PdfFrontMatterRenderMode.Value;
        if (source.PdfUseFrontMatterVisualTheme.HasValue) options.UseFrontMatterTheme = source.PdfUseFrontMatterVisualTheme.Value;
        if (source.PdfUseFrontMatterMetadata.HasValue) options.UseFrontMatterMetadata = source.PdfUseFrontMatterMetadata.Value;
        if (source.PdfUseFirstHeadingAsTitle.HasValue) options.UseFirstHeadingAsTitle = source.PdfUseFirstHeadingAsTitle.Value;
        if (source.PdfCreateOutlineFromHeadings.HasValue) options.CreateOutlineFromHeadings = source.PdfCreateOutlineFromHeadings.Value;

        return options;
    }

    internal static void SetPdfResultVariables(IMarkdownPdfOptionSource source, PSCmdlet command, OfficeIMO.Pdf.PdfSaveResult result)
    {
        PdfCommandUtilities.SetVariable(command, source.PdfWarningVariable, result.Warnings);
        PdfCommandUtilities.SetVariable(command, source.PdfConversionReportVariable, result.Report);
    }

    internal static MarkdownVisualTheme CreateTheme(OfficeVisualThemeKind kind) => MarkdownVisualTheme.Create(kind);

    private static MarkdownWriteOptions CreateWriteProfile(OfficeMarkdownWriteProfile? profile)
    {
        return profile switch
        {
            OfficeMarkdownWriteProfile.Portable => MarkdownWriteOptions.CreatePortableProfile(),
            OfficeMarkdownWriteProfile.HtmlImage => MarkdownWriteOptions.CreateHtmlImageProfile(),
            _ => MarkdownWriteOptions.CreateOfficeIMOProfile()
        };
    }

    private static bool HasReaderOverrides(IMarkdownReaderOptionSource source) =>
        !string.IsNullOrWhiteSpace(source.BaseUri)
        || source.MaxInputCharacters.HasValue
        || source.NormalizeInput.HasValue
        || source.DisallowFileUrls.HasValue
        || source.AllowDataUrls.HasValue
        || source.AllowMailtoUrls.HasValue
        || source.AllowProtocolRelativeUrls.HasValue
        || source.RestrictUrlSchemes.HasValue
        || source.AllowedUrlScheme is { Length: > 0 };

    private static string ResolveLineEnding(string value)
    {
        return value.Trim().ToUpperInvariant() switch
        {
            "CRLF" => "\r\n",
            "LF" => "\n",
            "CR" => "\r",
            "\\R\\N" => "\r\n",
            "\\N" => "\n",
            "\\R" => "\r",
            _ => value
        };
    }

}
