using System;
using System.Collections;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Markdown;

internal interface IMarkdownReaderOptionSource
{
    MarkdownLoadOptions? ReaderOptions { get; }
    MarkdownLoadOptions.MarkdownDialectProfile? Profile { get; }
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
    MarkdownSaveOptions? WriteOptions { get; }
    OfficeMarkdownWriteProfile? WriteProfile { get; }
    MarkdownImageRenderingMode? ImageRenderingMode { get; }
    string? LineEnding { get; }
    string? UnorderedListMarker { get; }
}

internal static class MarkdownOptionUtilities
{
    /// <summary>Captures native reader settings and applies explicit hashtable overrides to an independent snapshot.</summary>
    internal static MarkdownLoadOptions? BindReaderOptions(EngineIntrinsics engine, object[]? inputs)
    {
        MarkdownLoadOptions? options = null;
        var binder = new OfficeOptionsTransformationAttribute(typeof(MarkdownLoadOptions));
        foreach (object input in inputs ?? Array.Empty<object>())
        {
            object? value = OfficeArgumentValue.Unwrap(input);
            if (value is MarkdownLoadOptions typed)
            {
                if (options != null) throw new PSArgumentException("Supply one Markdown reader options object before any hashtable overrides.");
                options = typed.Clone();
            }
            else if (value is IDictionary dictionary)
            {
                options ??= MarkdownLoadOptions.CreateOfficeIMOProfile();
                options = ((MarkdownLoadOptions)binder.ApplyTo(engine, options, dictionary)).Clone();
            }
            else throw new PSArgumentException("Options accepts Markdown reader settings or a hashtable of native reader properties.");
        }
        return options;
    }

    internal static MarkdownLoadOptions? BuildReaderOptions(IMarkdownReaderOptionSource source)
    {
        if (source.ReaderOptions != null && source.Profile.HasValue)
        {
            throw new PSArgumentException("Specify either -ReaderOptions or -Profile, not both.");
        }

        if (source.ReaderOptions == null && !source.Profile.HasValue && !HasReaderOverrides(source))
        {
            return null;
        }

        var options = source.ReaderOptions?.Clone()
            ?? (source.Profile.HasValue
                ? MarkdownLoadOptions.CreateProfile(source.Profile.Value)
                : MarkdownLoadOptions.CreateOfficeIMOProfile());

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

    internal static MarkdownSaveOptions? BuildWriteOptions(IMarkdownWriteOptionSource source, MarkdownSaveOptions? defaults = null)
    {
        if (source.WriteOptions != null && source.WriteProfile.HasValue)
        {
            throw new PSArgumentException("Specify either -WriteOptions or -WriteProfile, not both.");
        }
        if (source.WriteOptions == null && !source.WriteProfile.HasValue && !source.ImageRenderingMode.HasValue
            && string.IsNullOrWhiteSpace(source.LineEnding) && string.IsNullOrWhiteSpace(source.UnorderedListMarker))
        {
            return null;
        }

        var options = source.WriteOptions?.Clone() ?? (source.WriteProfile.HasValue
            ? CreateWriteProfile(source.WriteProfile) : defaults?.Clone() ?? CreateWriteProfile(null));

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

    private static MarkdownSaveOptions CreateWriteProfile(OfficeMarkdownWriteProfile? profile)
    {
        return profile switch
        {
            OfficeMarkdownWriteProfile.Portable => MarkdownSaveOptions.CreatePortableProfile(),
            OfficeMarkdownWriteProfile.HtmlImage => MarkdownSaveOptions.CreateHtmlImageProfile(),
            _ => MarkdownSaveOptions.CreateOfficeIMOProfile()
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
