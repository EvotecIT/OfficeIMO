using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Restores streams carried by an RC4 CryptoAPI EncryptedSummary stream.</summary>
    internal static class LegacyPptEncryptedSummary {
        private const string EncryptedSummaryStream = "EncryptedSummary";
        private const string SummaryInformationStream = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStream =
            "\u0005DocumentSummaryInformation";

        internal static byte[] Encrypt(OfficeCompoundFile source,
            OfficeBinaryRc4CryptoApiSession session) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (session == null) throw new ArgumentNullException(nameof(session));

            byte[] summary = GetPropertyStreamOrEmpty(source,
                SummaryInformationStream,
                OfficeOlePropertySetWriter.SummaryInformationFormatId);
            byte[] documentSummary = GetPropertyStreamOrEmpty(source,
                DocumentSummaryInformationStream,
                OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId);
            return OfficeBinaryEncryptedSummary.Encrypt(new[] {
                new OfficeBinaryEncryptedSummary.Entry(SummaryInformationStream, true, summary),
                new OfficeBinaryEncryptedSummary.Entry(DocumentSummaryInformationStream, true, documentSummary)
            }, session);
        }

        internal static byte[] CreateEmptyDocumentSummaryInformation() =>
            OfficeOlePropertySetWriter.CreatePropertySet((
                OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId,
                OfficeOlePropertySetWriter.CreateSection(
                    Array.Empty<OfficeOleProperty>())));

        internal static IReadOnlyDictionary<string, byte[]> Decrypt(
            OfficeCompoundFile source,
            OfficeBinaryRc4CryptoApiSession session,
            LegacyPptImportOptions options,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(
                    nameof(decodedStorageBudget));
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (!source.Streams.TryGetValue(EncryptedSummaryStream,
                    out byte[]? encrypted)) {
                throw new InvalidDataException(
                    "The RC4 CryptoAPI encryption header requires an EncryptedSummary stream.");
            }
            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in OfficeBinaryEncryptedSummary.Decrypt(encrypted, session, long.MaxValue,
                (isStream, size) => { if (!isStream) decodedStorageBudget.Consume(size); }, cancellationToken)) {
                if (entry.IsStream) AddReplacement(replacements, entry.Name, entry.Bytes);
                else AddStorageReplacements(replacements, entry.Name, entry.Bytes,
                    options, decodedStorageBudget, cancellationToken);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (!replacements.ContainsKey(SummaryInformationStream)
                || !replacements.ContainsKey(DocumentSummaryInformationStream)) {
                throw new InvalidDataException(
                    "The EncryptedSummary stream must carry both Office document-property streams.");
            }
            return replacements;
        }

        private static byte[] GetPropertyStreamOrEmpty(
            OfficeCompoundFile source, string name, Guid formatId) {
            if (source.Streams.TryGetValue(name, out byte[]? bytes)) {
                return bytes;
            }
            return OfficeOlePropertySetWriter.CreatePropertySet((formatId,
                OfficeOlePropertySetWriter.CreateSection(
                    Array.Empty<OfficeOleProperty>())));
        }

        internal static void AddStorageReplacements(
            IDictionary<string, byte[]> replacements, string storageName,
            byte[] bytes, LegacyPptImportOptions options,
            LegacyPptDecodedStorageBudget decodedStorageBudget,
            CancellationToken cancellationToken = default) {
            if (replacements == null) {
                throw new ArgumentNullException(nameof(replacements));
            }
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (decodedStorageBudget == null) {
                throw new ArgumentNullException(
                    nameof(decodedStorageBudget));
            }
            cancellationToken.ThrowIfCancellationRequested();
            int maxDirectoryEntries = Math.Max(1,
                Math.Min(65536, options.MaxRecordCount));
            int maxStreamCount = Math.Max(1,
                Math.Min(32768, options.MaxRecordCount));
            var readOptions = new OfficeCompoundReadOptions(
                maxDirectoryEntries, maxStreamCount,
                maxStreamBytes: int.MaxValue,
                maxTotalStreamBytes: long.MaxValue,
                streamSizeValidator: (_, size) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    decodedStorageBudget.Consume(checked((int)size));
                }, preserveResourceLimits: true);
            if (!OfficeCompoundFileReader.TryRead(bytes, readOptions,
                    out OfficeCompoundFile? storage, out string? error)
                || storage == null) {
                decodedStorageBudget.ThrowIfExceeded();
                throw new InvalidDataException(
                    $"EncryptedSummary storage '{storageName}' is not a valid compound file: {error}");
            }
            int streamCount = 0;
            foreach (OfficeCompoundFileEntry entry in storage.Entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!entry.IsStream || entry.IsFallback) continue;
                AddReplacement(replacements,
                    storageName + "/" + entry.Path,
                    storage.Streams[entry.Path]);
                streamCount++;
            }
            if (streamCount == 0) {
                throw new InvalidDataException(
                    $"EncryptedSummary storage '{storageName}' contains no streams.");
            }
        }

        private static void AddReplacement(
            IDictionary<string, byte[]> replacements, string path,
            byte[] bytes) {
            if (replacements.ContainsKey(path)) {
                throw new InvalidDataException(
                    $"EncryptedSummary contains duplicate compound path '{path}'.");
            }
            replacements.Add(path, bytes);
        }

    }
}
