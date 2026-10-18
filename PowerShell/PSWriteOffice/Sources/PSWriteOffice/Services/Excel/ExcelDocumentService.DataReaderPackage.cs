using System;
using System.Data;
using System.IO;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static partial class ExcelDocumentService
{
    internal static ExcelDataSetImportResult WriteDataReaderPackage(
        string filePath,
        IDataReader reader,
        ExcelTabularWriteOptions options,
        bool overwrite)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        if (options == null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        var targetPath = Path.GetFullPath(filePath);
        if (!overwrite && File.Exists(targetPath))
        {
            throw new IOException($"File '{targetPath}' already exists.");
        }

        var directory = Path.GetDirectoryName(targetPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var temporaryPath = AtomicFileWriter.CreateTemporaryPath(
            directory ?? Directory.GetCurrentDirectory(),
            Path.GetFileName(targetPath));
        try
        {
            ExcelDataSetImportResult result;
            using (var stream = new FileStream(
                temporaryPath,
                FileMode.CreateNew,
                FileAccess.ReadWrite,
                FileShare.None,
                81920,
                FileOptions.SequentialScan))
            {
                result = ExcelDocument.WriteDataReader(stream, reader, options);
                stream.Flush();
            }

            CommitDataReaderPackage(temporaryPath, targetPath, overwrite);
            temporaryPath = string.Empty;
            return result;
        }
        finally
        {
            if (!string.IsNullOrEmpty(temporaryPath) && File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }

    private static void CommitDataReaderPackage(string temporaryPath, string targetPath, bool overwrite)
    {
        if (!File.Exists(targetPath))
        {
            File.Move(temporaryPath, targetPath);
            return;
        }

        if (!overwrite)
        {
            throw new IOException($"File '{targetPath}' already exists.");
        }

#if FRAMEWORK
        File.Replace(temporaryPath, targetPath, null);
#else
        File.Move(temporaryPath, targetPath, overwrite: true);
#endif
    }
}
