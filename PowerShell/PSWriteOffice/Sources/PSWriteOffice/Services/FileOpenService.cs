using System;
using System.Diagnostics;
using System.IO;

namespace PSWriteOffice.Services;

internal static class FileOpenService
{
    public static void Open(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = resolvedPath,
            UseShellExecute = true
        });
    }
}
