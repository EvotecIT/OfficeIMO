using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using System.Security.Cryptography;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>Adapts the XLS FilePass prefix and BIFF record mask to the shared CryptoAPI session.</summary>
    internal static class BiffRc4CryptoApiEncryption {
        internal static bool IsRc4CryptoApiFilePass(BiffRecord record) =>
            record.Type == (ushort)BiffRecordType.FilePass && record.Payload.Length >= 6
            && BiffRecordReader.ReadUInt16(record.Payload, 0) == 1
            && BiffRecordReader.ReadUInt16(record.Payload, 2) >= 2
            && BiffRecordReader.ReadUInt16(record.Payload, 2) <= 4
            && BiffRecordReader.ReadUInt16(record.Payload, 4) == 2;

        internal static OfficeBinaryRc4CryptoApiSession OpenSession(BiffRecord filePassRecord, string password) {
            byte[] payload = filePassRecord.Payload;
            if (!IsRc4CryptoApiFilePass(filePassRecord))
                throw new InvalidDataException("The FilePass record is not RC4 CryptoAPI.");
            var header = new byte[payload.Length - 2];
            Buffer.BlockCopy(payload, 2, header, 0, header.Length);
            return OfficeBinaryRc4CryptoApiSession.Open(header, password);
        }

        internal static bool TryDecrypt(byte[] workbookStream, BiffRecord filePassRecord,
            string password, List<LegacyXlsImportDiagnostic> diagnostics, out byte[] decryptedWorkbookStream) {
            decryptedWorkbookStream = workbookStream;
            try {
                var session = OpenSession(filePassRecord, password);
                decryptedWorkbookStream = BiffRc4RecordStreamCipher.Transform(workbookStream, session);
                return true;
            } catch (CryptographicException) {
                AddError(diagnostics, filePassRecord, "XLS-BIFF-FILEPASS-PASSWORD-INVALID",
                    "The supplied legacy XLS password did not match the workbook FilePass verifier.");
            } catch (Exception error) when (error is InvalidDataException or NotSupportedException or ArgumentException or OverflowException) {
                AddError(diagnostics, filePassRecord, "XLS-BIFF-FILEPASS-INVALID",
                    "The RC4 CryptoAPI FilePass descriptor could not be read. " + error.Message);
            }
            return false;
        }

        private static void AddError(List<LegacyXlsImportDiagnostic> diagnostics, BiffRecord record, string code, string message) =>
            diagnostics.Add(new LegacyXlsImportDiagnostic(LegacyXlsDiagnosticSeverity.Error, code, message,
                recordOffset: record.Offset, recordType: record.Type, detailCode: "Encryption:FilePass:Rc4CryptoApi"));
    }
}
