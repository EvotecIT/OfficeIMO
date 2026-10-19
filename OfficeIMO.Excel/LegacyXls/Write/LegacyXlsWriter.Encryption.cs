using OfficeIMO.Excel.LegacyXls.Biff;
using System.Security.Cryptography;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    internal static void ValidateEncryptionPassword(string password) {
        if (password == null) throw new ArgumentNullException(nameof(password));
        if (password.Length < 1 || password.Length > 15) throw new ArgumentException("Classic XLS encryption requires 1–15 UTF-16 password code units.", nameof(password));
    }

    internal static byte[] WriteEncryptedWorkbook(ExcelDocument document, string password, long maximumBytes) {
        ValidateEncryptionPassword(password);
        if (document.LegacyXlsSourceCompoundFile?.Streams.Keys.Any(name => name.StartsWith("_SX_DB_CUR/", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "Revision Log", StringComparison.OrdinalIgnoreCase)) == true)
            throw new NotSupportedException("Encrypted XLS writing does not yet support retained pivot-cache or revision BIFF streams. They cannot be copied as plaintext into an encrypted workbook.");
        var salt = new byte[16]; var verifier = new byte[16];
        using (var random = RandomNumberGenerator.Create()) { random.GetBytes(salt); random.GetBytes(verifier); }
        byte[] filePass = BiffRc4Encryption.CreateFilePassPayload(password, salt, verifier);
        if (document.TryGetLegacyXlsPreservationPlan(out var plan)) return plan!.WriteEncrypted(filePass, password, salt, maximumBytes);
        byte[] workbook = BuildWorkbookStream(document, filePass, maximumBytes);
        return BuildCompoundFile(document, BiffRc4Encryption.EncryptWorkbookStream(workbook, password, salt), maximumBytes);
    }
}
