using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private byte[] CreateSignatureValidationSnapshot(WordSignatureValidationOptions options) {
            if (_ownedPackageStream == null) {
                throw new InvalidDataException("The current OPC package has no encoded package stream available for validation.");
            }
            if (_ownedPackageStream.Length > options.MaxPackageBytes) {
                throw new InvalidDataException("The current OPC package exceeds the " + options.MaxPackageBytes + " byte validation limit.");
            }
            if (_wordprocessingDocument.FileOpenAccess == FileAccess.Read) {
                return _ownedPackageStream.ToArray();
            }

            byte[] encodedPackage = _ownedPackageStream.ToArray();
            List<OpenXmlPart> sourceParts = EnumerateSignatureSnapshotParts(
                    _wordprocessingDocument,
                    options.MaxPackageParts)
                .OrderByDescending(part => part.IsRootElementLoaded)
                .ToList();
            List<DataPart> sourceDataParts = _wordprocessingDocument.DataParts.ToList();
            EnsureSignatureSnapshotPartCount(sourceParts.Count, sourceDataParts.Count, options.MaxPackageParts);
            foreach (OpenXmlPart sourcePart in sourceParts) {
                EnsureSignatureSnapshotPartWithinLimit(sourcePart, options.MaxPartBytes);
            }
            foreach (DataPart sourceDataPart in sourceDataParts) {
                EnsureSignatureSnapshotDataPartWithinLimit(sourceDataPart, options.MaxPartBytes);
            }

            using var snapshot = new SignatureValidationSnapshotMemoryStream(options.MaxPackageBytes);
            snapshot.Write(encodedPackage, 0, encodedPackage.Length);
            var currentPartPayloads = new Dictionary<Uri, byte[]>();
            var currentDataPartPayloads = sourceDataParts.ToDictionary(
                part => part.Uri,
                part => ReadCurrentSignatureSnapshotDataPart(part, options.MaxPartBytes));
            HashSet<Uri> encodedPartUris;
            using (var encodedStream = new MemoryStream(encodedPackage, writable: false)) {
                using WordprocessingDocument encodedPackageDocument = WordprocessingDocument.Open(encodedStream, false);
                Dictionary<Uri, OpenXmlPart> encodedParts = EnumerateSignatureSnapshotParts(
                        encodedPackageDocument,
                        options.MaxPackageParts)
                    .ToDictionary(part => part.Uri);
                List<DataPart> encodedDataParts = encodedPackageDocument.DataParts.ToList();
                EnsureSignatureSnapshotPartCount(encodedParts.Count, encodedDataParts.Count, options.MaxPackageParts);
                encodedPartUris = new HashSet<Uri>(encodedParts.Keys.Concat(encodedDataParts.Select(part => part.Uri)));
                foreach (OpenXmlPart sourcePart in sourceParts) {
                    OpenXmlPartRootElement? sourceRoot = sourcePart.IsRootElementLoaded
                        ? sourcePart.RootElement
                        : null;
                    if (sourceRoot != null &&
                        encodedParts.TryGetValue(sourcePart.Uri, out OpenXmlPart? encodedPart) &&
                        encodedPart.RootElement is OpenXmlPartRootElement encodedRoot &&
                        AreSignatureSnapshotRootsEquivalent(sourceRoot, encodedRoot)) {
                        continue;
                    }
                    currentPartPayloads[sourcePart.Uri] = ReadCurrentSignatureSnapshotPart(
                        sourcePart,
                        options.MaxPartBytes);
                }
            }
            ApplyCurrentSignatureSnapshotState(
                snapshot,
                sourceParts,
                sourceDataParts,
                encodedPartUris,
                currentPartPayloads,
                currentDataPartPayloads,
                options.MaxPackageParts,
                options.MaxPartBytes);
            if (snapshot.Length > options.MaxPackageBytes) {
                throw new InvalidDataException("The current OPC package exceeds the " + options.MaxPackageBytes + " byte validation limit.");
            }
            return snapshot.ToArray();
        }

        private static void EnsureSignatureSnapshotPartCount(int openXmlPartCount, int dataPartCount, int maxPackageParts) {
            if (openXmlPartCount > maxPackageParts - dataPartCount) {
                throw new SignatureValidationSnapshotResourceException(
                    "The OPC package contains more than " + maxPackageParts +
                    " parts while creating the current-state validation snapshot.");
            }
        }

        private static void EnsureSignatureSnapshotPartWithinLimit(OpenXmlPart part, long maxPartBytes) {
            using (Stream input = part.GetStream(FileMode.Open, FileAccess.Read)) {
                if (input.CanSeek && input.Length > maxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        (part.IsRootElementLoaded ? "A pending package part " : "The current package part ") +
                        part.Uri + " exceeds the " +
                        maxPartBytes + " byte validation limit.");
                }
                if (!input.CanSeek) CopySignatureSnapshotPart(input, Stream.Null, maxPartBytes, part.Uri);
            }

            if (!part.IsRootElementLoaded || part.RootElement == null) return;
            using var boundedOutput = new SignatureValidationPartWriteStream(Stream.Null, maxPartBytes);
            part.RootElement.Save(boundedOutput);
        }

        private static byte[] ReadCurrentSignatureSnapshotPart(OpenXmlPart part, long maxPartBytes) {
            using var output = new MemoryStream();
            if (part.IsRootElementLoaded && part.RootElement != null) {
                using var boundedOutput = new SignatureValidationPartWriteStream(output, maxPartBytes);
                part.RootElement.Save(boundedOutput);
            } else {
                using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
                CopySignatureSnapshotPart(input, output, maxPartBytes, part.Uri);
            }
            return output.ToArray();
        }

        private static void EnsureSignatureSnapshotDataPartWithinLimit(DataPart part, long maxPartBytes) {
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            if (input.CanSeek && input.Length > maxPartBytes) {
                throw new SignatureValidationSnapshotResourceException(
                    "The current package data part " + part.Uri + " exceeds the " +
                    maxPartBytes + " byte validation limit.");
            }
            if (!input.CanSeek) CopySignatureSnapshotPart(input, Stream.Null, maxPartBytes, part.Uri);
        }

        private static byte[] ReadCurrentSignatureSnapshotDataPart(DataPart part, long maxPartBytes) {
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            CopySignatureSnapshotPart(input, output, maxPartBytes, part.Uri);
            return output.ToArray();
        }

        private void ApplyCurrentSignatureSnapshotState(
            SignatureValidationSnapshotMemoryStream snapshot,
            IReadOnlyList<OpenXmlPart> currentParts,
            IReadOnlyList<DataPart> currentDataParts,
            HashSet<Uri> encodedPartUris,
            IReadOnlyDictionary<Uri, byte[]> currentPartPayloads,
            IReadOnlyDictionary<Uri, byte[]> currentDataPartPayloads,
            int maxPackageParts,
            long maxPartBytes) {
            snapshot.Position = 0;
            using var snapshotArchive = new ZipArchive(snapshot, ZipArchiveMode.Update, leaveOpen: true);
            var currentPartUris = new HashSet<Uri>(currentParts.Select(part => part.Uri)
                .Concat(currentDataParts.Select(part => part.Uri)));
            foreach (Uri removedPartUri in encodedPartUris.Where(uri => !currentPartUris.Contains(uri))) {
                snapshotArchive.GetEntry(GetSignatureSnapshotEntryName(removedPartUri))?.Delete();
            }
            foreach (KeyValuePair<Uri, byte[]> payload in currentPartPayloads) {
                ReplaceSignatureSnapshotEntry(snapshotArchive, GetSignatureSnapshotEntryName(payload.Key), payload.Value);
            }
            foreach (KeyValuePair<Uri, byte[]> payload in currentDataPartPayloads) {
                ReplaceSignatureSnapshotEntry(snapshotArchive, GetSignatureSnapshotEntryName(payload.Key), payload.Value);
            }
            var currentContentTypes = currentParts
                .Select(part => (part.Uri, part.ContentType))
                .Concat(currentDataParts.Select(part => (part.Uri, part.ContentType)))
                .ToDictionary(item => item.Uri, item => item.ContentType);
            UpdateSignatureSnapshotContentTypes(snapshotArchive, currentContentTypes, encodedPartUris, maxPartBytes);

            Dictionary<string, byte[]> currentRelationships = BuildCurrentSignatureSnapshotRelationships(currentParts, maxPackageParts, maxPartBytes);
            foreach (KeyValuePair<string, byte[]> relationship in currentRelationships) {
                ZipArchiveEntry? existingEntry = snapshotArchive.GetEntry(relationship.Key);
                if (existingEntry != null) {
                    byte[] existingRelationships = ReadSignatureSnapshotEntry(existingEntry, maxPartBytes);
                    if (AreSignatureSnapshotRelationshipsEquivalent(
                        existingRelationships,
                        relationship.Value,
                        relationship.Key,
                        maxPartBytes)) {
                        continue;
                    }
                }
                ReplaceSignatureSnapshotEntry(snapshotArchive, relationship.Key, relationship.Value);
            }
            foreach (OpenXmlPart currentPart in currentParts) {
                string relationshipEntryName = GetSignatureSnapshotRelationshipEntryName(currentPart.Uri);
                if (!currentRelationships.ContainsKey(relationshipEntryName)) {
                    snapshotArchive.GetEntry(relationshipEntryName)?.Delete();
                }
            }
            if (!currentRelationships.ContainsKey("_rels/.rels")) {
                snapshotArchive.GetEntry("_rels/.rels")?.Delete();
            }
        }

        private Dictionary<string, byte[]> BuildCurrentSignatureSnapshotRelationships(
            IReadOnlyList<OpenXmlPart> currentParts,
            int maxPackageParts,
            long maxPartBytes) {
            var relationships = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            AddCurrentSignatureSnapshotRelationships(
                relationships,
                "_rels/.rels",
                _wordprocessingDocument,
                maxPartBytes);
            int visited = 0;
            foreach (OpenXmlPart part in currentParts) {
                if (++visited > maxPackageParts) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The OPC package contains more than " + maxPackageParts +
                        " parts during relationship snapshot creation.");
                }
                AddCurrentSignatureSnapshotRelationships(
                    relationships,
                    GetSignatureSnapshotRelationshipEntryName(part.Uri),
                    part,
                    maxPartBytes);
            }
            return relationships;
        }

        private static void AddCurrentSignatureSnapshotRelationships(
            IDictionary<string, byte[]> destination,
            string entryName,
            OpenXmlPartContainer container,
            long maxPartBytes) {
            XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            var values = new List<(string Id, string Type, string Target, bool External)>();
            values.AddRange(container.Parts.Select(pair => (
                pair.RelationshipId,
                pair.OpenXmlPart.RelationshipType,
                GetRelativeSignatureSnapshotRelationshipTarget(entryName, pair.OpenXmlPart.Uri),
                false)));
            values.AddRange(container.ExternalRelationships.Select(relationship => (
                relationship.Id,
                relationship.RelationshipType,
                relationship.Uri.ToString(),
                true)));
            values.AddRange(container.HyperlinkRelationships.Select(relationship => (
                relationship.Id,
                relationship.RelationshipType,
                relationship.IsExternal
                    ? relationship.Uri.ToString()
                    : GetRelativeSignatureSnapshotRelationshipTarget(entryName, relationship.Uri),
                relationship.IsExternal)));
            values.AddRange(container.DataPartReferenceRelationships.Select(relationship => (
                relationship.Id,
                relationship.RelationshipType,
                relationship.IsExternal
                    ? relationship.Uri.ToString()
                    : GetRelativeSignatureSnapshotRelationshipTarget(entryName, relationship.Uri),
                relationship.IsExternal)));
            if (values.Count == 0) return;

            var root = new XElement(packageRelationships + "Relationships",
                values.OrderBy(value => value.Id, StringComparer.Ordinal).Select(value =>
                    new XElement(packageRelationships + "Relationship",
                        new XAttribute("Id", value.Id),
                        new XAttribute("Type", value.Type),
                        new XAttribute("Target", value.Target),
                        value.External ? new XAttribute("TargetMode", "External") : null)));
            destination[entryName] = SerializeSignatureSnapshotXml(root, maxPartBytes);
        }

        private static string GetRelativeSignatureSnapshotRelationshipTarget(string relationshipEntryName, Uri targetUri) {
            string sourcePath = GetSignatureSnapshotRelationshipSourcePath(relationshipEntryName);
            var sourceUri = new Uri("http://officeimo.invalid/" + sourcePath, UriKind.Absolute);
            var absoluteTarget = new Uri("http://officeimo.invalid/" + targetUri.ToString().TrimStart('/'), UriKind.Absolute);
            return Uri.UnescapeDataString(sourceUri.MakeRelativeUri(absoluteTarget).ToString());
        }

        private static void UpdateSignatureSnapshotContentTypes(
            ZipArchive archive,
            IReadOnlyDictionary<Uri, string> currentContentTypes,
            HashSet<Uri> encodedPartUris,
            long maxPartBytes) {
            const string entryName = "[Content_Types].xml";
            ZipArchiveEntry? entry = archive.GetEntry(entryName);
            if (entry == null) {
                throw new InvalidDataException("The OPC package has no [Content_Types].xml part.");
            }
            byte[] bytes = ReadSignatureSnapshotEntry(entry, maxPartBytes);
            using var input = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = maxPartBytes,
                XmlResolver = null
            });
            XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            XElement root = document.Root ?? throw new InvalidDataException("The OPC content-type catalog is empty.");
            XNamespace contentTypes = root.Name.Namespace;
            var currentPartUris = new HashSet<Uri>(currentContentTypes.Keys);
            foreach (XElement overrideElement in root.Elements(contentTypes + "Override").ToList()) {
                string? partName = overrideElement.Attribute("PartName")?.Value;
                if (partName != null && Uri.TryCreate(partName, UriKind.Relative, out Uri? uri) &&
                    encodedPartUris.Contains(uri) && !currentPartUris.Contains(uri)) {
                    overrideElement.Remove();
                }
            }
            foreach (KeyValuePair<Uri, string> part in currentContentTypes) {
                XElement? overrideElement = root.Elements(contentTypes + "Override")
                    .FirstOrDefault(element => string.Equals(
                        element.Attribute("PartName")?.Value,
                        part.Key.ToString(),
                        StringComparison.OrdinalIgnoreCase));
                if (overrideElement == null) {
                    root.Add(new XElement(contentTypes + "Override",
                        new XAttribute("PartName", part.Key.ToString()),
                        new XAttribute("ContentType", part.Value)));
                } else {
                    overrideElement.SetAttributeValue("ContentType", part.Value);
                }
            }
            ReplaceSignatureSnapshotEntry(archive, entryName, SerializeSignatureSnapshotXml(root, maxPartBytes));
        }

        private static byte[] SerializeSignatureSnapshotXml(XElement root, long maxPartBytes) {
            using var output = new MemoryStream();
            using (var boundedOutput = new SignatureValidationPartWriteStream(output, maxPartBytes)) {
                using XmlWriter writer = XmlWriter.Create(boundedOutput, new XmlWriterSettings {
                    Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                    OmitXmlDeclaration = false,
                    CloseOutput = false
                });
                root.Save(writer);
            }
            return output.ToArray();
        }

        private static void ReplaceSignatureSnapshotEntry(ZipArchive archive, string entryName, byte[] payload) {
            archive.GetEntry(entryName)?.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream output = replacement.Open();
            output.Write(payload, 0, payload.Length);
        }

        private static string GetSignatureSnapshotEntryName(Uri partUri) =>
            partUri.ToString().TrimStart('/');

        private static string GetSignatureSnapshotRelationshipEntryName(Uri partUri) {
            string entryName = GetSignatureSnapshotEntryName(partUri);
            int slashIndex = entryName.LastIndexOf('/');
            string folder = slashIndex < 0 ? string.Empty : entryName.Substring(0, slashIndex + 1);
            string fileName = slashIndex < 0 ? entryName : entryName.Substring(slashIndex + 1);
            return folder + "_rels/" + fileName + ".rels";
        }

        private static byte[] ReadSignatureSnapshotEntry(ZipArchiveEntry entry, long maxPartBytes) {
            if (entry.Length > maxPartBytes) {
                throw new SignatureValidationSnapshotResourceException(
                    "The current package relationship part /" + entry.FullName + " exceeds the " +
                    maxPartBytes + " byte validation limit.");
            }
            using Stream input = entry.Open();
            using var output = new MemoryStream();
            CopySignatureSnapshotPart(input, output, maxPartBytes, new Uri("/" + entry.FullName, UriKind.Relative));
            return output.ToArray();
        }

        private static bool AreSignatureSnapshotRelationshipsEquivalent(
            byte[] encodedRelationships,
            byte[] snapshotRelationships,
            string relationshipEntryName,
            long maxPartBytes) {
            IReadOnlyList<string> encoded = ReadSignatureSnapshotRelationshipIdentities(
                encodedRelationships,
                relationshipEntryName,
                maxPartBytes);
            IReadOnlyList<string> current = ReadSignatureSnapshotRelationshipIdentities(
                snapshotRelationships,
                relationshipEntryName,
                maxPartBytes);
            return encoded.SequenceEqual(current, StringComparer.Ordinal);
        }

        private static IReadOnlyList<string> ReadSignatureSnapshotRelationshipIdentities(
            byte[] bytes,
            string relationshipEntryName,
            long maxPartBytes) {
            using var input = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = maxPartBytes,
                XmlResolver = null
            });
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            XElement? root = document.Root;
            if (root == null || !string.Equals(root.Name.LocalName, "Relationships", StringComparison.Ordinal)) {
                return new[] { System.Convert.ToBase64String(bytes) };
            }
            return root.Elements()
                .Select(element => element.Name.ToString() + "|" + string.Join("|", element.Attributes()
                    .Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => attribute.Name + "=" + NormalizeSignatureSnapshotRelationshipAttribute(
                        element,
                        attribute,
                        relationshipEntryName))
                    .OrderBy(value => value, StringComparer.Ordinal)))
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
        }

        private static string NormalizeSignatureSnapshotRelationshipAttribute(
            XElement relationship,
            XAttribute attribute,
            string relationshipEntryName) {
            if (!string.Equals(attribute.Name.LocalName, "Target", StringComparison.Ordinal) ||
                string.Equals(
                    relationship.Attributes().FirstOrDefault(item => string.Equals(item.Name.LocalName, "TargetMode", StringComparison.Ordinal))?.Value,
                    "External",
                    StringComparison.OrdinalIgnoreCase)) {
                return attribute.Value;
            }
            string sourcePartPath = GetSignatureSnapshotRelationshipSourcePath(relationshipEntryName);
            var sourceUri = new Uri("http://officeimo.invalid/" + sourcePartPath, UriKind.Absolute);
            return new Uri(sourceUri, attribute.Value).PathAndQuery;
        }

        private static string GetSignatureSnapshotRelationshipSourcePath(string relationshipEntryName) {
            if (string.Equals(relationshipEntryName, "_rels/.rels", StringComparison.OrdinalIgnoreCase)) {
                return string.Empty;
            }
            int markerIndex = relationshipEntryName.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
            if (markerIndex < 0) return relationshipEntryName;
            string folder = relationshipEntryName.Substring(0, markerIndex + 1);
            string fileName = relationshipEntryName.Substring(markerIndex + "/_rels/".Length);
            return folder + fileName.Substring(0, fileName.Length - ".rels".Length);
        }

        private static void CopySignatureSnapshotPart(
            Stream input,
            Stream output,
            long maxPartBytes,
            Uri partUri) {
            byte[] buffer = new byte[81920];
            long copied = 0;
            while (true) {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                copied = checked(copied + read);
                if (copied > maxPartBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current package part " + partUri + " exceeds the " +
                        maxPartBytes + " byte validation limit.");
                }
                output.Write(buffer, 0, read);
            }
        }

        private sealed class SignatureValidationSnapshotMemoryStream : MemoryStream {
            private readonly long _maxBytes;

            internal SignatureValidationSnapshotMemoryStream(long maxBytes) {
                _maxBytes = maxBytes;
            }

            public override void Write(byte[] buffer, int offset, int count) {
                EnsureWithinLimit(Math.Max(Length, checked(Position + count)));
                base.Write(buffer, offset, count);
            }

            public override void WriteByte(byte value) {
                EnsureWithinLimit(Math.Max(Length, checked(Position + 1)));
                base.WriteByte(value);
            }

            public override void SetLength(long value) {
                EnsureWithinLimit(value);
                base.SetLength(value);
            }

            private void EnsureWithinLimit(long value) {
                if (value > _maxBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The current OPC package exceeds the " + _maxBytes +
                        " byte validation-snapshot limit.");
                }
            }
        }

        private sealed class SignatureValidationPartWriteStream : Stream {
            private readonly Stream _inner;
            private readonly long _maxBytes;
            private long _written;

            internal SignatureValidationPartWriteStream(Stream inner, long maxBytes) {
                _inner = inner;
                _maxBytes = maxBytes;
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _written;
            public override long Position { get => _written; set => throw new NotSupportedException(); }
            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) {
                long next = checked(_written + count);
                if (next > _maxBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "A pending package part exceeds the " + _maxBytes +
                        " byte validation-snapshot serialization limit.");
                }
                _inner.Write(buffer, offset, count);
                _written = next;
            }

            public override void WriteByte(byte value) {
                if (_written >= _maxBytes) {
                    throw new SignatureValidationSnapshotResourceException(
                        "A pending package part exceeds the " + _maxBytes +
                        " byte validation-snapshot serialization limit.");
                }
                _inner.WriteByte(value);
                _written++;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateSignatureSnapshotParts(
            OpenXmlPartContainer container,
            int maxPackageParts) {
            var pending = new Stack<OpenXmlPart>(container.Parts.Select(pair => pair.OpenXmlPart));
            var visited = new HashSet<Uri>();
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part.Uri)) continue;
                if (visited.Count > maxPackageParts) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The OPC package contains more than " + maxPackageParts + " parts during validation snapshot creation.");
                }
                yield return part;
                foreach (IdPartPair child in part.Parts) pending.Push(child.OpenXmlPart);
            }
        }

        private static bool AreSignatureSnapshotRootsEquivalent(OpenXmlElement source, OpenXmlElement snapshot) {
            var pending = new Stack<(OpenXmlElement Source, OpenXmlElement Snapshot)>();
            pending.Push((source, snapshot));
            while (pending.Count > 0) {
                (OpenXmlElement sourceElement, OpenXmlElement snapshotElement) = pending.Pop();
                if (!string.Equals(sourceElement.LocalName, snapshotElement.LocalName, StringComparison.Ordinal) ||
                    !string.Equals(sourceElement.NamespaceUri, snapshotElement.NamespaceUri, StringComparison.Ordinal)) {
                    return false;
                }
                IList<OpenXmlAttribute> sourceAttributes = sourceElement.GetAttributes();
                IList<OpenXmlAttribute> snapshotAttributes = snapshotElement.GetAttributes();
                if (sourceAttributes.Count != snapshotAttributes.Count || sourceAttributes.Any(sourceAttribute =>
                    !snapshotAttributes.Any(snapshotAttribute =>
                        string.Equals(sourceAttribute.LocalName, snapshotAttribute.LocalName, StringComparison.Ordinal) &&
                        string.Equals(sourceAttribute.NamespaceUri, snapshotAttribute.NamespaceUri, StringComparison.Ordinal) &&
                        string.Equals(sourceAttribute.Value, snapshotAttribute.Value, StringComparison.Ordinal)))) {
                    return false;
                }
                List<KeyValuePair<string, string>> sourceNamespaces = sourceElement.NamespaceDeclarations.ToList();
                List<KeyValuePair<string, string>> snapshotNamespaces = snapshotElement.NamespaceDeclarations.ToList();
                if (sourceNamespaces.Count != snapshotNamespaces.Count || sourceNamespaces.Any(sourceNamespace =>
                    !snapshotNamespaces.Any(snapshotNamespace =>
                        string.Equals(sourceNamespace.Key, snapshotNamespace.Key, StringComparison.Ordinal) &&
                        string.Equals(sourceNamespace.Value, snapshotNamespace.Value, StringComparison.Ordinal)))) {
                    return false;
                }
                if (sourceElement.ChildElements.Count != snapshotElement.ChildElements.Count) return false;
                if (sourceElement.ChildElements.Count == 0 && !string.Equals(
                    sourceElement.InnerText,
                    snapshotElement.InnerText,
                    StringComparison.Ordinal)) {
                    return false;
                }
                for (int index = sourceElement.ChildElements.Count - 1; index >= 0; index--) {
                    pending.Push((sourceElement.ChildElements[index], snapshotElement.ChildElements[index]));
                }
            }
            return true;
        }

        private sealed class SignatureValidationSnapshotResourceException : Exception {
            internal SignatureValidationSnapshotResourceException(string message) : base(message) { }
        }

    }
}
