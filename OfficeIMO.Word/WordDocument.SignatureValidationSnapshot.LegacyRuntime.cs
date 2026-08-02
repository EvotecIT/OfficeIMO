using System.IO.Compression;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static bool RequiresLegacyRuntimeSignatureSnapshot() =>
            RuntimeInformation.FrameworkDescription.StartsWith(".NET Framework", StringComparison.OrdinalIgnoreCase);

        // System.IO.Packaging on .NET Framework invalidates ProgressiveCrcCalculatingStream after
        // in-memory package edits. Clone is the supported way to materialize that live state without
        // mutating the source package or rereading its stale part streams.
#pragma warning disable OOXML0001
        private byte[] CreateLegacyRuntimeSignatureValidationSnapshot(WordSignatureValidationOptions options) {
            IPackage sourcePackage = _wordprocessingDocument.GetPackage();
            List<IPackagePart> sourceParts = GetBoundedSignatureSnapshotPackageParts(
                sourcePackage,
                options.MaxPackageParts,
                "creating");
            foreach (IPackagePart sourcePart in sourceParts) {
                using Stream source = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
                if (source.CanSeek && source.Length > options.MaxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + sourcePart.Uri + " exceeds the " +
                        options.MaxPartBytes + " byte validation limit.");
                }
            }

            byte[] encodedPackage = _ownedPackageStream!.ToArray();
            using var snapshot = new SignatureValidationSnapshotMemoryStream(options.MaxPackageBytes);
            using (_wordprocessingDocument.Clone(snapshot, true)) { }
            byte[] currentPackage = snapshot.ToArray();
            HashSet<Uri> unchangedSerializedRoots = FindUnchangedSignatureSnapshotRoots(
                currentPackage,
                encodedPackage,
                options.MaxPackageParts,
                options.MaxPartBytes);
            HashSet<string> unchangedRelationships = FindUnchangedSignatureSnapshotRelationships(
                currentPackage,
                encodedPackage,
                options.MaxPartBytes);
            if (unchangedSerializedRoots.Count > 0 || unchangedRelationships.Count > 0) {
                snapshot.Position = 0;
                using (var snapshotArchive = new ZipArchive(snapshot, ZipArchiveMode.Update, leaveOpen: true)) {
                    using var restorationArchive = new ZipArchive(
                        new MemoryStream(encodedPackage, writable: false),
                        ZipArchiveMode.Read);
                    foreach (Uri unchangedRootUri in unchangedSerializedRoots) {
                        RestoreSignatureSnapshotEntry(
                            snapshotArchive,
                            restorationArchive,
                            GetSignatureSnapshotEntryName(unchangedRootUri),
                            options.MaxPartBytes);
                    }
                    foreach (string relationshipEntryName in unchangedRelationships) {
                        RestoreSignatureSnapshotEntry(
                            snapshotArchive,
                            restorationArchive,
                            relationshipEntryName,
                            options.MaxPartBytes);
                    }
                }
                currentPackage = snapshot.ToArray();
            }
            EnforceLegacySignatureSnapshotPayloadBudgets(currentPackage, encodedPackage, options);
            return currentPackage;
        }
#pragma warning restore OOXML0001

        private static void RestoreSignatureSnapshotEntry(
            ZipArchive snapshotArchive,
            ZipArchive encodedArchive,
            string entryName,
            long maxPartBytes) {
            ZipArchiveEntry? encodedEntry = encodedArchive.GetEntry(entryName);
            if (encodedEntry == null) return;
            ReplaceSignatureSnapshotEntry(
                snapshotArchive,
                entryName,
                ReadSignatureSnapshotEntry(encodedEntry, maxPartBytes));
        }

        private static HashSet<Uri> FindUnchangedSignatureSnapshotRoots(
            byte[] currentPackage,
            byte[] encodedPackage,
            int maxPackageParts,
            long maxPartBytes) {
            using var currentStream = new MemoryStream(currentPackage, writable: false);
            using var encodedStream = new MemoryStream(encodedPackage, writable: false);
            using WordprocessingDocument currentDocument = WordprocessingDocument.Open(currentStream, false);
            using WordprocessingDocument encodedDocument = WordprocessingDocument.Open(encodedStream, false);
            Dictionary<Uri, OpenXmlPart> currentParts = EnumerateSignatureSnapshotParts(
                    currentDocument,
                    maxPackageParts)
                .ToDictionary(part => part.Uri);
            Dictionary<Uri, OpenXmlPart> encodedParts = EnumerateSignatureSnapshotParts(
                    encodedDocument,
                    maxPackageParts)
                .ToDictionary(part => part.Uri);
            var unchanged = new HashSet<Uri>();
            foreach (KeyValuePair<Uri, OpenXmlPart> currentPart in currentParts) {
                if (!encodedParts.TryGetValue(currentPart.Key, out OpenXmlPart? encodedPart) ||
                    currentPart.Value.RootElement is not OpenXmlPartRootElement currentRoot ||
                    encodedPart.RootElement is not OpenXmlPartRootElement encodedRoot) {
                    continue;
                }
                byte[] currentPayload = SerializeSignatureSnapshotRoot(currentRoot, maxPartBytes);
                byte[] encodedPayload = SerializeSignatureSnapshotRoot(encodedRoot, maxPartBytes);
                if (currentPayload.SequenceEqual(encodedPayload)) unchanged.Add(currentPart.Key);
            }
            return unchanged;
        }

        private static HashSet<string> FindUnchangedSignatureSnapshotRelationships(
            byte[] currentPackage,
            byte[] encodedPackage,
            long maxPartBytes) {
            using var currentArchive = new ZipArchive(
                new MemoryStream(currentPackage, writable: false),
                ZipArchiveMode.Read);
            using var encodedArchive = new ZipArchive(
                new MemoryStream(encodedPackage, writable: false),
                ZipArchiveMode.Read);
            var unchanged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry currentEntry in currentArchive.Entries.Where(entry =>
                         entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))) {
                ZipArchiveEntry? encodedEntry = encodedArchive.GetEntry(currentEntry.FullName);
                if (encodedEntry == null) continue;
                byte[] currentRelationships = ReadSignatureSnapshotEntry(currentEntry, maxPartBytes);
                byte[] encodedRelationships = ReadSignatureSnapshotEntry(encodedEntry, maxPartBytes);
                if (AreSignatureSnapshotRelationshipsEquivalent(
                    encodedRelationships,
                    currentRelationships,
                    currentEntry.FullName,
                    maxPartBytes)) {
                    unchanged.Add(currentEntry.FullName);
                }
            }
            return unchanged;
        }

#pragma warning disable OOXML0001
        private static void EnforceLegacySignatureSnapshotPayloadBudgets(
            byte[] currentPackage,
            byte[] encodedPackage,
            WordSignatureValidationOptions options) {
            long changedPayloadBytes = 0;
            using var currentStream = new MemoryStream(currentPackage, writable: false);
            using WordprocessingDocument currentDocument = WordprocessingDocument.Open(currentStream, false);
            List<IPackagePart> currentParts = GetBoundedSignatureSnapshotPackageParts(
                currentDocument.GetPackage(),
                options.MaxPackageParts,
                "verifying");
            using var encodedArchive = new ZipArchive(
                new MemoryStream(encodedPackage, writable: false),
                ZipArchiveMode.Read);
            foreach (IPackagePart currentPart in currentParts) {
                ZipArchiveEntry? encodedEntry = encodedArchive.GetEntry(GetSignatureSnapshotEntryName(currentPart.Uri));
                using Stream currentPartStream = currentPart.GetStream(FileMode.Open, FileAccess.Read);
                long currentLength;
                bool unchanged;
                if (encodedEntry == null) {
                    currentLength = CopySignatureSnapshotPart(
                        currentPartStream,
                        Stream.Null,
                        options.MaxPartBytes,
                        currentPart.Uri);
                    unchanged = false;
                } else {
                    using Stream encodedPartStream = encodedEntry.Open();
                    unchanged = CompareSignatureSnapshotParts(
                        currentPartStream,
                        encodedPartStream,
                        options.MaxPartBytes,
                        currentPart.Uri,
                        out currentLength);
                }
                if (!unchanged) {
                    ReserveSignatureSnapshotPayload(
                        ref changedPayloadBytes,
                        currentLength,
                        options.MaxTotalDigestBytes,
                        currentPart.Uri);
                }
            }
        }
#pragma warning restore OOXML0001

#pragma warning disable OOXML0001
        private static List<IPackagePart> GetBoundedSignatureSnapshotPackageParts(
            IPackage package,
            int maxPackageParts,
            string operation) {
            var parts = new List<IPackagePart>(Math.Min(maxPackageParts, 256));
            foreach (IPackagePart part in package.GetParts()) {
                if (parts.Count >= maxPackageParts) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The OPC package contains more than " + maxPackageParts +
                        " parts while " + operation + " the current-state validation snapshot.");
                }
                parts.Add(part);
            }
            return parts;
        }
#pragma warning restore OOXML0001

        private static bool CompareSignatureSnapshotParts(
            Stream current,
            Stream encoded,
            long maxPartBytes,
            Uri partUri,
            out long currentLength) {
            byte[] currentBuffer = new byte[81920];
            byte[] encodedBuffer = new byte[81920];
            currentLength = 0;
            bool equal = true;
            while (true) {
                int currentRead = current.Read(currentBuffer, 0, currentBuffer.Length);
                int encodedRead = encoded.Read(encodedBuffer, 0, encodedBuffer.Length);
                if (currentRead == 0) return equal && encodedRead == 0;
                currentLength = checked(currentLength + currentRead);
                if (currentLength > maxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + partUri + " exceeds the " +
                        maxPartBytes + " byte validation limit.");
                }
                if (currentRead != encodedRead || !currentBuffer.AsSpan(0, currentRead).SequenceEqual(
                    encodedBuffer.AsSpan(0, encodedRead))) {
                    equal = false;
                }
            }
        }
    }
}
