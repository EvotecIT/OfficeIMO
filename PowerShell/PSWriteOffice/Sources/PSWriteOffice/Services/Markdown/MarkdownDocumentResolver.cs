using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Services.Markdown;

internal static class MarkdownDocumentResolver
{
    public static MarkdownDocument Resolve(
        PSCmdlet cmdlet,
        string parameterSetName,
        string documentParameterSetName,
        MarkdownDocument? document,
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

            return MarkdownDocument.Load(resolvedPath, effectiveOptions);
        }

        return MarkdownDocument.Parse(text ?? string.Empty, effectiveOptions);
    }

    public static MarkdownDocument Resolve(
        PSCmdlet cmdlet,
        string parameterSetName,
        string documentParameterSetName,
        MarkdownDocument? document,
        string? inputPath,
        string? text,
        MarkdownLoadOptions? options,
        MarkdownLoadOptions.MarkdownDialectProfile? profile)
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
            ? MarkdownLoadOptions.CreateProfile(profile.Value)
            : null);

        if (!string.IsNullOrEmpty(inputPath))
        {
            var resolvedPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            return MarkdownDocument.Load(resolvedPath, effectiveOptions);
        }

        return MarkdownDocument.Parse(text ?? string.Empty, effectiveOptions);
    }
}
