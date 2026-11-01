using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Visio;

namespace PSWriteOffice.Services.Visio;

internal static class VisioCommandUtilities
{
    internal static string ResolvePath(PSCmdlet cmdlet, string path)
        => OfficeFileCommandUtilities.ResolvePath(cmdlet, path);

    internal static void EnsureDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
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

}
