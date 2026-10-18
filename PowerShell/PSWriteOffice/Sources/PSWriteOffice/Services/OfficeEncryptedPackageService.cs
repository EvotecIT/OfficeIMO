using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace PSWriteOffice.Services;

internal static class OfficeEncryptedPackageService
{
    internal static bool HasCompoundFileSignature(string path)
    {
        if (!File.Exists(path))
        {
            return false;
        }

        byte[] signature = new byte[8];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        if (stream.Read(signature, 0, signature.Length) != signature.Length)
        {
            return false;
        }

        return signature[0] == 0xD0 && signature[1] == 0xCF &&
            signature[2] == 0x11 && signature[3] == 0xE0 &&
            signature[4] == 0xA1 && signature[5] == 0xB1 &&
            signature[6] == 0x1A && signature[7] == 0xE1;
    }

    public static ExcelDocument LoadExcel(string path, string password, bool readOnly, bool autoSave)
    {
        if (autoSave)
        {
            throw new NotSupportedException("Encrypted Excel workbooks require explicit Save-OfficeExcel -Password or Close-OfficeExcel -Save -Password. OfficeIMO does not support SaveOnDispose for encrypted sources.");
        }

        return ExcelDocument.LoadEncrypted(path, password, new ExcelLoadOptions
        {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SaveExcel(ExcelDocument document, string path, string password, bool openExcel, ExcelSaveOptions? saveOptions)
    {
        document.SaveEncrypted(path, password, saveOptions);
        if (openExcel)
        {
            FileOpenService.Open(path);
        }
    }

    public static WordDocument LoadWord(string path, string password, bool readOnly, bool autoSave)
    {
        if (autoSave)
        {
            throw new NotSupportedException("Encrypted Word documents require explicit Save-OfficeWord -Password or Close-OfficeWord -Save -Password. OfficeIMO does not support SaveOnDispose for encrypted sources.");
        }

        return WordDocument.LoadEncrypted(path, password, new WordLoadOptions
        {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SaveWord(WordDocument document, string path, string password, bool openWord)
    {
        document.SaveEncrypted(path, password);
        if (openWord)
        {
            FileOpenService.Open(path);
        }
    }

    public static PowerPointPresentation OpenPowerPoint(string path, string password, bool readOnly = false)
    {
        return PowerPointPresentation.LoadEncrypted(path, password, new PowerPointLoadOptions
        {
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
    }

    public static void SavePowerPoint(PowerPointPresentation presentation, string path, string password)
    {
        presentation.SaveEncrypted(path, password);
    }
}
