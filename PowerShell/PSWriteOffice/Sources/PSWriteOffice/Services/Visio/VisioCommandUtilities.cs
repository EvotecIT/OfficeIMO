using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

namespace PSWriteOffice.Services.Visio;

internal static class VisioCommandUtilities
{
    internal static string ResolvePath(PSCmdlet cmdlet, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Path cannot be empty.", nameof(path));
        }

        var providerPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(cmdlet.SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    internal static void EnsureDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    internal static string ResolveImageOutputPath(
        PSCmdlet cmdlet,
        string path,
        OfficeImageExportFormat format)
    {
        var resolvedPath = ResolvePath(cmdlet, path);
        var extension = Path.GetExtension(resolvedPath);
        if (string.IsNullOrEmpty(extension))
        {
            return resolvedPath + format.GetFileExtension();
        }

        if (!format.HasFileExtension(extension))
        {
            throw new PSArgumentException(
                $"Output path extension '{extension}' does not match the selected {format} format. " +
                $"Use {format.GetFileExtension()} or omit the extension.",
                nameof(path));
        }

        return resolvedPath;
    }

    internal static VisioDocument ResolveDocument(PSCmdlet cmdlet, VisioDocument? document, string? path)
    {
        if (document != null)
        {
            return document;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Provide -Document or -Path.", nameof(path));
        }

        return VisioDocument.Load(ResolvePath(cmdlet, path!));
    }

    internal static VisioImageExportOptions BuildImageOptions(
        PSCmdlet cmdlet,
        int pageIndex,
        double? pixelsPerInch,
        string? backgroundColor,
        bool transparent,
        bool noText,
        string? fontFilePath,
        string? fontFaceName,
        int? fontCollectionIndex,
        bool noStencilArtwork,
        bool noConnectorLabels,
        bool noConnectorLabelOverlapResolution,
        int? supersampling,
        bool includeSvgXmlDeclaration)
    {
        var options = new VisioImageExportOptions
        {
            PageIndex = pageIndex,
            RenderText = !noText,
            FontFilePath = string.IsNullOrWhiteSpace(fontFilePath) ? null : ResolvePath(cmdlet, fontFilePath!),
            FontFaceName = fontFaceName,
            FontCollectionIndex = fontCollectionIndex,
            RenderStencilArtwork = !noStencilArtwork,
            RenderConnectorLabels = !noConnectorLabels,
            ResolveConnectorLabelOverlaps = !noConnectorLabelOverlapResolution,
            IncludeSvgXmlDeclaration = includeSvgXmlDeclaration
        };

        if (pixelsPerInch.HasValue)
        {
            options.TargetDpi = pixelsPerInch.Value;
        }

        if (supersampling.HasValue)
        {
            options.Supersampling = supersampling.Value;
        }

        options.BackgroundColor = ResolveBackgroundColor(backgroundColor, transparent);
        return options;
    }

    private static OfficeColor ResolveBackgroundColor(string? color, bool transparent)
    {
        if (transparent)
        {
            return OfficeColor.Transparent;
        }

        return string.IsNullOrWhiteSpace(color) ? OfficeColor.White : OfficeColor.Parse(color!);
    }
}
