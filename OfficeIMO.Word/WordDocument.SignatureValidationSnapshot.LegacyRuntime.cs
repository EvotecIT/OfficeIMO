using System.IO.Compression;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static bool RequiresLegacyRuntimeSignatureSnapshot() =>
            RuntimeInformation.FrameworkDescription.StartsWith(".NET Framework", StringComparison.OrdinalIgnoreCase);

        private static byte[]? CaptureLegacyValidationEncodedPackage(byte[] sourceBytes, bool readOnly) =>
            !readOnly && RequiresLegacyRuntimeSignatureSnapshot()
                ? (byte[])sourceBytes.Clone()
                : null;

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
            Dictionary<Uri, byte[]> pendingRootPayloads = CaptureLegacyPendingRootPayloads(options);
            Dictionary<string, byte[]> currentRelationships = BuildLegacySignatureSnapshotRelationships(
                sourcePackage,
                sourceParts,
                options.MaxPartBytes);
            foreach (IPackagePart sourcePart in sourceParts) {
                using Stream source = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
                if (source.CanSeek && source.Length > options.MaxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + sourcePart.Uri + " exceeds the " +
                        options.MaxPartBytes + " byte validation limit.");
                }
            }

            byte[] encodedPackage = _legacyValidationEncodedPackageBytes ?? _ownedPackageStream!.ToArray();
            if (encodedPackage.LongLength > options.MaxPackageBytes) {
                throw new SignatureValidationSnapshotResourceException(
                    "The encoded OPC package exceeds the " + options.MaxPackageBytes +
                    " byte validation-snapshot limit.");
            }
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
                pendingRootPayloads,
                encodedPackage,
                options.MaxPackageParts,
                options.MaxPartBytes);
            currentPackage = ApplyLegacySignatureSnapshotState(
                currentPackage,
                encodedPackage,
                unchangedSerializedRoots,
                currentRelationships,
                options.MaxPartBytes,
                options.MaxPackageBytes);
            EnforceLegacySignatureSnapshotPayloadBudgets(currentPackage, encodedPackage, options);
            return currentPackage;
        }
#pragma warning restore OOXML0001

        private MemoryStream DetachLegacyValidationLivePackageStream(byte[] encodedPackage) {
            if (_legacyValidationLivePackageStream != null) return _legacyValidationLivePackageStream;
            MemoryStream livePackageStream = _ownedPackageStream!;
            _legacyValidationLivePackageStream = livePackageStream;
            _ownedPackageStream = new MemoryStream(encodedPackage, writable: false);
            _legacyValidationEncodedPackageBytes = null;
            return livePackageStream;
        }

        private Dictionary<Uri, byte[]> CaptureLegacyPendingRootPayloads(WordSignatureValidationOptions options) {
            var payloads = new Dictionary<Uri, byte[]>();
            MainDocumentPart? mainPart = _wordprocessingDocument.MainDocumentPart;
            OpenXmlPartRootElement? mainRoot = mainPart?.IsRootElementLoaded == true
                ? mainPart.RootElement
                : null;
            if (mainPart != null && mainRoot != null) {
                payloads[mainPart.Uri] = SerializeSignatureSnapshotRoot(mainRoot, options.MaxPartBytes);
            }
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
                    payloads[part.Uri] = SerializeSignatureSnapshotRoot(root, options.MaxPartBytes);
                }
            }
            return payloads;
        }

#pragma warning disable OOXML0001
        private static Dictionary<string, byte[]> BuildLegacySignatureSnapshotRelationships(
            IPackage package,
            IReadOnlyList<IPackagePart> parts,
            long maxPartBytes) {
            var relationships = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            AddLegacySignatureSnapshotRelationships(
                relationships,
                "_rels/.rels",
                package.Relationships,
                maxPartBytes);
            foreach (IPackagePart part in parts) {
                if (part.Uri.OriginalString.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                AddLegacySignatureSnapshotRelationships(
                    relationships,
                    GetSignatureSnapshotRelationshipEntryName(part.Uri),
                    part.Relationships,
                    maxPartBytes);
            }
            return relationships;
        }

        private static void AddLegacySignatureSnapshotRelationships(
            IDictionary<string, byte[]> destination,
            string entryName,
            IRelationshipCollection relationships,
            long maxPartBytes) {
            XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            IPackageRelationship[] values = relationships
                .OrderBy(relationship => relationship.Id, StringComparer.Ordinal)
                .ToArray();
            if (values.Length == 0) return;
            var root = new XElement(packageRelationships + "Relationships",
                values.Select(relationship =>
                    new XElement(packageRelationships + "Relationship",
                        new XAttribute("Id", relationship.Id),
                        new XAttribute("Type", relationship.RelationshipType),
                        new XAttribute("Target", relationship.TargetUri.ToString()),
                        relationship.TargetMode == TargetMode.External
                            ? new XAttribute("TargetMode", "External")
                            : null)));
            destination[entryName] = SerializeSignatureSnapshotXml(root, maxPartBytes);
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
            IReadOnlyDictionary<Uri, byte[]> currentRootPayloads,
            byte[] encodedPackage,
            int maxPackageParts,
            long maxPartBytes) {
            using var encodedStream = new MemoryStream(encodedPackage, writable: false);
            using WordprocessingDocument encodedDocument = WordprocessingDocument.Open(encodedStream, false);
            Dictionary<Uri, OpenXmlPart> encodedParts = EnumerateSignatureSnapshotParts(
                    encodedDocument,
                    maxPackageParts)
                .ToDictionary(part => part.Uri);
            var unchanged = new HashSet<Uri>();
            foreach (KeyValuePair<Uri, byte[]> currentRoot in currentRootPayloads) {
                if (!encodedParts.TryGetValue(currentRoot.Key, out OpenXmlPart? encodedPart) ||
                    encodedPart.RootElement is not OpenXmlPartRootElement encodedRoot) {
                    continue;
                }
                byte[] encodedPayload = SerializeSignatureSnapshotRoot(encodedRoot, maxPartBytes);
                if (currentRoot.Value.SequenceEqual(encodedPayload)) unchanged.Add(currentRoot.Key);
            }
            return unchanged;
        }

        private static byte[] ApplyLegacySignatureSnapshotState(
            byte[] currentPackage,
            byte[] encodedPackage,
            IReadOnlyCollection<Uri> unchangedSerializedRoots,
            IReadOnlyDictionary<string, byte[]> currentRelationships,
            long maxPartBytes,
            long maxPackageBytes) {
            using var snapshot = new SignatureValidationSnapshotMemoryStream(maxPackageBytes);
            snapshot.Write(currentPackage, 0, currentPackage.Length);
            snapshot.Position = 0;
            using (var snapshotArchive = new ZipArchive(snapshot, ZipArchiveMode.Update, leaveOpen: true)) {
                using var encodedArchive = new ZipArchive(
                    new MemoryStream(encodedPackage, writable: false),
                    ZipArchiveMode.Read);
                foreach (Uri unchangedRootUri in unchangedSerializedRoots) {
                    RestoreSignatureSnapshotEntry(
                        snapshotArchive,
                        encodedArchive,
                        GetSignatureSnapshotEntryName(unchangedRootUri),
                        maxPartBytes);
                }
                foreach (ZipArchiveEntry staleRelationships in snapshotArchive.Entries
                             .Where(entry => entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) &&
                                             !currentRelationships.ContainsKey(entry.FullName))
                             .ToList()) {
                    staleRelationships.Delete();
                }
                foreach (KeyValuePair<string, byte[]> relationship in currentRelationships) {
                    byte[] payload = relationship.Value;
                    ZipArchiveEntry? encodedEntry = encodedArchive.GetEntry(relationship.Key);
                    if (encodedEntry != null) {
                        byte[] encodedPayload = ReadSignatureSnapshotEntry(encodedEntry, maxPartBytes);
                        if (AreSignatureSnapshotRelationshipsEquivalent(
                            encodedPayload,
                            payload,
                            relationship.Key,
                            maxPartBytes)) {
                            payload = encodedPayload;
                        }
                    }
                    ReplaceSignatureSnapshotEntry(snapshotArchive, relationship.Key, payload);
                }
            }
            return snapshot.ToArray();
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
                if (currentRead != encodedRead || !BuffersEqual(currentBuffer, encodedBuffer, currentRead)) {
                    equal = false;
                }
            }
        }

        private static bool BuffersEqual(byte[] left, byte[] right, int count) {
            for (int index = 0; index < count; index++) {
                if (left[index] != right[index]) return false;
            }
            return true;
        }
    }
}
