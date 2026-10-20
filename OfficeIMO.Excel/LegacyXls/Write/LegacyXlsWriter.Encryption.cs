using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Core.Internal;
using System.Security.Cryptography;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    internal static void ValidateEncryptionPassword(string password) {
        if (password == null) throw new ArgumentNullException(nameof(password));
        if (password.Length < 1 || password.Length > 15) throw new ArgumentException("Classic XLS encryption requires 1–15 UTF-16 password code units.", nameof(password));
    }

    internal static byte[] WriteEncryptedWorkbook(ExcelDocument document, string password, long maximumBytes, ExcelNativeCalculationPolicy calculation = default, ExcelSaveOptions? options = null) {
        ValidateEncryptionPassword(password);
        if (document.LegacyXlsEncryptedSummaryEntries.Count != 0)
            return WriteCryptoApiWorkbook(document, password, maximumBytes, calculation, options);
        var salt = new byte[16]; var verifier = new byte[16];
        using (var random = RandomNumberGenerator.Create()) { random.GetBytes(salt); random.GetBytes(verifier); }
        byte[] filePass = BiffRc4Encryption.CreateFilePassPayload(password, salt, verifier);
        byte[] Encrypt(byte[] bytes) => BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt);
        if (document.TryGetLegacyXlsPreservationPlan(options, out var plan)) return plan!.WriteEncrypted(filePass, Encrypt, maximumBytes, calculation);
        byte[] workbook = BuildWorkbookStream(document, filePass, maximumBytes);
        return BuildCompoundFile(document, BiffRc4Encryption.EncryptWorkbookStream(workbook, password, salt), maximumBytes,
            bytes => BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt));
    }

    /// <summary>Retains protected auxiliary carriers instead of exposing them through a classic-RC4 format downgrade.</summary>
    private static byte[] WriteCryptoApiWorkbook(ExcelDocument document, string password, long maximumBytes,
        ExcelNativeCalculationPolicy calculation, ExcelSaveOptions? options) {
        var protectedEntries = document.LegacyXlsEncryptedSummaryEntries;
        var session = OfficeBinaryRc4CryptoApiSession.Create(password, 128,
            protectedEntries.ContainsKey("\u0005SummaryInformation"), out byte[] header);
        var filePass = new byte[header.Length + 2]; filePass[0] = 1;
        Buffer.BlockCopy(header, 0, filePass, 2, header.Length);
        byte[] Encrypt(byte[] bytes) => BiffRc4RecordStreamCipher.Transform(bytes, session);
        byte[] prepared = document.TryGetLegacyXlsPreservationPlan(options, out var plan)
            ? plan!.WriteEncrypted(filePass, Encrypt, maximumBytes, calculation)
            : BuildCompoundFile(document, Encrypt(BuildWorkbookStream(document, filePass, maximumBytes)), maximumBytes, Encrypt);
        if (!OfficeCompoundFileReader.TryRead(prepared, out OfficeCompoundFile? compound, out string? error))
            throw new InvalidDataException("Encrypted XLS output could not be prepared: " + error);
        return LegacyXlsCompoundEncryption.ProtectSummaryEntries(compound!, protectedEntries, session, maximumBytes);
    }
}
