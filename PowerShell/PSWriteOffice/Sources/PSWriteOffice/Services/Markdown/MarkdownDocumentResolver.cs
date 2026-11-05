using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Markdown;

internal static class MarkdownDocumentResolver
{
    internal static MarkdownDocument Resolve(PSCmdlet cmdlet, string parameterSetName,
        MarkdownDocument? document, string path, string literalPath, string text, object[]? options, bool captureSyntax = false) {
        if (parameterSetName == "Document")
            return document ?? OfficeDocumentContext.Find<MarkdownDocument>() ?? MarkdownDslContext.Require(cmdlet).Document;
        var engine = (EngineIntrinsics)cmdlet.SessionState.PSVariable.GetValue("ExecutionContext");
        var reader = MarkdownOptionUtilities.BindReaderOptions(engine, options);
        if (captureSyntax) { reader ??= new MarkdownLoadOptions(); reader.CaptureSyntaxTree = true; }
        if (parameterSetName == "Text") return MarkdownDocument.Parse(text, reader);
        var sources = OfficePath.ResolveInput(cmdlet.SessionState, parameterSetName == "LiteralPath" ? literalPath : path,
            parameterSetName == "LiteralPath");
        if (sources.Count != 1) throw new PSArgumentException("The inspection source must resolve to one Markdown file; pipe files to inspect each separately.");
        return MarkdownDocument.Load(sources[0], reader);
    }

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
