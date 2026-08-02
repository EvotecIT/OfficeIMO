using System.IO.Compression;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static bool RequiresLegacyRuntimeSignatureSnapshot() =>
            RuntimeInformation.FrameworkDescription.StartsWith(".NET Framework", StringComparison.OrdinalIgnoreCase);

        // System.IO.Packaging on .NET Framework invalidates ProgressiveCrcCalculatingStream after
        // in-memory package edits. Flush the live package once, retain that stream for the document,
        // and keep the original encoded bytes in the validation baseline instead of rereading stale
        // part streams through Clone.
#pragma warning disable OOXML0001
        private byte[] CreateLegacyRuntimeSignatureValidationSnapshot(WordSignatureValidationOptions options) {
            IPackage sourcePackage = _wordprocessingDocument.GetPackage();
            List<IPackagePart> sourceParts = GetBoundedSignatureSnapshotPackageParts(
                sourcePackage,
                options.MaxPackageParts,
                "creating");
            EnforceLegacyPendingRootSerializationBudgets(options);
            foreach (IPackagePart sourcePart in sourceParts) {
                using Stream source = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
                if (source.CanSeek && source.Length > options.MaxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + sourcePart.Uri + " exceeds the " +
                        options.MaxPartBytes + " byte validation limit.");
                }
            }

            byte[] encodedPackage = _ownedPackageStream!.ToArray();
            MemoryStream livePackageStream = DetachLegacyValidationLivePackageStream(encodedPackage);
            _wordprocessingDocument.Save();
            sourcePackage.Save();
            livePackageStream.Flush();
            if (livePackageStream.Length > options.MaxPackageBytes) {
                throw new SignatureValidationSnapshotResourceException(
                    "The current OPC package exceeds the " + options.MaxPackageBytes +
                    " byte validation-snapshot limit.");
            }
            byte[] currentPackage = livePackageStream.ToArray();
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
                using var snapshot = new SignatureValidationSnapshotMemoryStream(options.MaxPackageBytes);
                snapshot.Write(currentPackage, 0, currentPackage.Length);
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

        private MemoryStream DetachLegacyValidationLivePackageStream(byte[] encodedPackage) {
            if (_legacyValidationLivePackageStream != null) return _legacyValidationLivePackageStream;
            MemoryStream livePackageStream = _ownedPackageStream!;
            _legacyValidationLivePackageStream = livePackageStream;
            _ownedPackageStream = new MemoryStream(encodedPackage, writable: false);
            return livePackageStream;
        }

        private void EnforceLegacyPendingRootSerializationBudgets(WordSignatureValidationOptions options) {
            MainDocumentPart? mainPart = _wordprocessingDocument.MainDocumentPart;
            OpenXmlPartRootElement? mainRoot = mainPart?.IsRootElementLoaded == true
                ? mainPart.RootElement
                : null;
            if (mainRoot != null) SerializeSignatureSnapshotRoot(mainRoot, options.MaxPartBytes);
            List<OpenXmlPart> reachableParts;
            try {
                reachableParts = EnumerateSignatureSnapshotParts(
                    _wordprocessingDocument,
                    options.MaxPackageParts).ToList();
            } catch (InvalidOperationException) {
                // A pending part removal can leave a stale SDK relationship edge on .NET Framework.
                // The package-level part enumeration above remains authoritative for current parts.
                reachableParts = new List<OpenXmlPart>();
            }
            foreach (OpenXmlPart part in reachableParts) {
                if (part.IsRootElementLoaded &&
                    part.RootElement is OpenXmlPartRootElement root &&
                    !ReferenceEquals(root, mainRoot)) {
                    SerializeSignatureSnapshotRoot(root, options.MaxPartBytes);
                }
            }
        }

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
            using var currentArchive = new ZipArchive(
                new MemoryStream(currentPackage, writable: false), ZipArchiveMode.Read);
            using var encodedArchive = new ZipArchive(
                new MemoryStream(encodedPackage, writable: false), ZipArchiveMode.Read);
            foreach (ZipArchiveEntry currentEntry in currentArchive.Entries.Where(entry =>
                         !string.IsNullOrEmpty(entry.Name))) {
                Uri currentUri = new Uri("/" + currentEntry.FullName, UriKind.Relative);
                if (currentEntry.Length > options.MaxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + currentUri + " exceeds the " +
                        options.MaxPartBytes + " byte validation limit.");
                }
                ZipArchiveEntry? encodedEntry = encodedArchive.GetEntry(currentEntry.FullName);
                bool unchanged = false;
                if (encodedEntry != null) {
                    using Stream currentPartStream = currentEntry.Open();
                    using Stream encodedPartStream = encodedEntry.Open();
                    unchanged = CompareSignatureSnapshotParts(
                        currentPartStream,
                        encodedPartStream,
                        options.MaxPartBytes,
                        currentUri,
                        out _);
                }
                if (!unchanged) {
                    ReserveSignatureSnapshotPayload(
                        ref changedPayloadBytes,
                        currentEntry.Length,
                        options.MaxTotalDigestBytes,
                        currentUri);
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
