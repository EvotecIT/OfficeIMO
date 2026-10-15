using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Keeps canonical native ODF admission failures terminal during loading and projection.</summary>
internal static class OdfReaderAdmission {
    internal static T Run<T>(Func<T> read) {
        try { return read(); }
        catch (OfficeInspectionResourceLimitException exception) {
            throw new ReaderResourceLimitException(exception.Owner + "." + exception.LimitName, exception.Maximum, exception);
        } catch (OfficePackageSecurityException exception) when (exception.Rule == OfficePackageSecurityRule.CompressionRatio) {
            throw new ReaderAdmissionLimitException("OpenDocument.MaxCompressionRatio", exception.Limit, exception.ObservedValue, exception);
        } catch (OdfEncryptedPackageException exception) when (exception.Reason == OdfEncryptionFailureReason.ResourceLimitExceeded) {
            throw new ReaderAdmissionLimitException("OpenDocument.Encryption.ResourceLimitExceeded", null, null, exception);
        }
    }

    internal static IEnumerable<T> Enumerate<T>(Func<IEnumerable<T>> read) {
        using IEnumerator<T> enumerator = Run(() => read().GetEnumerator());
        while (Run(enumerator.MoveNext)) yield return Run(() => enumerator.Current);
    }
}
