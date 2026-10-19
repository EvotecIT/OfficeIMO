using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Services.Markdown;

internal static class MarkdownDocumentResolver
{
    public static MarkdownDoc Resolve(
        PSCmdlet cmdlet,
        string parameterSetName,
        string documentParameterSetName,
        MarkdownDoc? document,
        string? inputPath,
        string? text,
        IMarkdownReaderOptionSource optionSource)
    {
        if (string.Equals(parameterSetName, documentParameterSetName, StringComparison.Ordinal))
        {
            return document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        var effectiveOptions = MarkdownOptionUtilities.BuildReaderOptions(optionSource);

        if (!string.IsNullOrEmpty(inputPath))
        {
            var resolvedPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            return MarkdownDoc.Load(resolvedPath, effectiveOptions);
        }

        return MarkdownReader.Parse(text ?? string.Empty, effectiveOptions);
    }

    public static MarkdownDoc Resolve(
        PSCmdlet cmdlet,
        string parameterSetName,
        string documentParameterSetName,
        MarkdownDoc? document,
        string? inputPath,
        string? text,
        MarkdownReaderOptions? options,
        MarkdownReaderOptions.MarkdownDialectProfile? profile)
    {
        if (options != null && profile.HasValue)
        {
            throw new PSArgumentException("Specify either -Options or -Profile, not both.");
        }

        if (string.Equals(parameterSetName, documentParameterSetName, StringComparison.Ordinal))
        {
            return document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        var effectiveOptions = options ?? (profile.HasValue
            ? MarkdownReaderOptions.CreateProfile(profile.Value)
            : null);

        if (!string.IsNullOrEmpty(inputPath))
        {
            var resolvedPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            return MarkdownDoc.Load(resolvedPath, effectiveOptions);
        }

        return MarkdownReader.Parse(text ?? string.Empty, effectiveOptions);
    }
}
