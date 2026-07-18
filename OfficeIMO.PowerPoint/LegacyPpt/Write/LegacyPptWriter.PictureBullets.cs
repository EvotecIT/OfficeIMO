using System.Collections.ObjectModel;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordBlipCollection9ForWrite = 0x07F8;
        private const ushort RecordBlipEntity9AtomForWrite = 0x07F9;
        private const int MaximumPictureBulletBytes = 64 * 1024 * 1024;

        internal static bool TryRewriteDocumentPictureBullets(
            LegacyPptRecord document,
            LegacyPptWriterPictureBulletCatalog catalog,
            bool replaceExisting, out byte[] bytes) {
            if (document == null) throw new ArgumentNullException(
                nameof(document));
            if (catalog == null) throw new ArgumentNullException(
                nameof(catalog));
            bytes = document.CopyRecordBytes();
            if (document.Version != 0x0F) return false;
            LegacyPptRecord[] infoLists = document.Children.Where(record =>
                record.Type == RecordDocumentInfoList).ToArray();
            if (infoLists.Length > 1) return false;
            byte[]? collection = catalog.BuildCollectionRecord();
            byte[]? rewrittenInfo = null;
            if (infoLists.Length == 1) {
                if (!TryRewritePictureBulletInfoList(infoLists[0],
                        collection, replaceExisting,
                        out rewrittenInfo)) return false;
            } else if (collection != null) {
                rewrittenInfo = BuildContainer(RecordDocumentInfoList,
                    instance: 0, new[] {
                        BuildContainer(RecordProgTags, instance: 0,
                            new[] {
                                BuildPictureBulletPpt9BinaryTag(collection)
                            })
                    });
            }
            if (rewrittenInfo == null) return true;

            var children = new List<byte[]>(document.Children.Count
                + (infoLists.Length == 0 ? 1 : 0));
            bool inserted = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (infoLists.Length == 1
                    && ReferenceEquals(child, infoLists[0])) {
                    children.Add(rewrittenInfo);
                    inserted = true;
                    continue;
                }
                if (infoLists.Length == 0 && !inserted
                    && (child.Type == RecordHeadersFooters
                        || child.Type == RecordSlideListWithText
                            && child.Instance != 1
                        || child.Type == RecordEndDocumentAtom)) {
                    children.Add(rewrittenInfo);
                    inserted = true;
                }
                children.Add(child.CopyRecordBytes());
            }
            if (!inserted) children.Add(rewrittenInfo);
            bytes = BuildRecord(document.Version, document.Instance,
                document.Type, Concat(children));
            return true;
        }

        private static bool TryRewritePictureBulletInfoList(
            LegacyPptRecord infoList, byte[]? collection,
            bool replaceExisting, out byte[] bytes) {
            bytes = infoList.CopyRecordBytes();
            if (infoList.Version != 0x0F || infoList.Instance != 0) {
                return false;
            }
            LegacyPptRecord[] tagLists = infoList.Children.Where(record =>
                record.Type == RecordProgTags).ToArray();
            if (tagLists.Length > 1) return false;
            byte[]? rewrittenTags = null;
            if (tagLists.Length == 1) {
                if (!TryRewritePictureBulletProgTags(tagLists[0],
                        collection, replaceExisting,
                        out rewrittenTags)) return false;
            } else if (collection != null) {
                rewrittenTags = BuildContainer(RecordProgTags, instance: 0,
                    new[] {
                        BuildPictureBulletPpt9BinaryTag(collection)
                    });
            }
            if (rewrittenTags == null) return true;
            var children = new List<byte[]>(infoList.Children.Count
                + (tagLists.Length == 0 ? 1 : 0));
            foreach (LegacyPptRecord child in infoList.Children) {
                children.Add(tagLists.Length == 1
                    && ReferenceEquals(child, tagLists[0])
                        ? rewrittenTags
                        : child.CopyRecordBytes());
            }
            if (tagLists.Length == 0) children.Add(rewrittenTags);
            bytes = BuildRecord(infoList.Version, infoList.Instance,
                infoList.Type, Concat(children));
            return true;
        }

        private static bool TryRewritePictureBulletProgTags(
            LegacyPptRecord progTags, byte[]? collection,
            bool replaceExisting, out byte[] bytes) {
            bytes = progTags.CopyRecordBytes();
            if (progTags.Version != 0x0F || progTags.Instance != 0) {
                return false;
            }
            LegacyPptRecord[] ppt9Tags = progTags.Children
                .Where(IsPpt9BinaryTag).ToArray();
            if (ppt9Tags.Length > 1) return false;
            byte[]? rewrittenPpt9 = null;
            if (ppt9Tags.Length == 1) {
                if (!TryRewritePictureBulletBinaryTag(ppt9Tags[0],
                        collection, replaceExisting,
                        out rewrittenPpt9)) return false;
            } else if (collection != null) {
                rewrittenPpt9 = BuildPictureBulletPpt9BinaryTag(collection);
            }
            if (rewrittenPpt9 == null) return true;
            var children = new List<byte[]>(progTags.Children.Count
                + (ppt9Tags.Length == 0 ? 1 : 0));
            foreach (LegacyPptRecord child in progTags.Children) {
                children.Add(ppt9Tags.Length == 1
                    && ReferenceEquals(child, ppt9Tags[0])
                        ? rewrittenPpt9
                        : child.CopyRecordBytes());
            }
            if (ppt9Tags.Length == 0) children.Add(rewrittenPpt9);
            bytes = BuildRecord(progTags.Version, progTags.Instance,
                progTags.Type, Concat(children));
            return true;
        }

        private static bool TryRewritePictureBulletBinaryTag(
            LegacyPptRecord binaryTag, byte[]? collection,
            bool replaceExisting, out byte[] bytes) {
            bytes = binaryTag.CopyRecordBytes();
            LegacyPptRecord[] blobs = binaryTag.Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (binaryTag.Version != 0x0F || binaryTag.Instance != 0
                || blobs.Length != 1 || blobs[0].Version != 0
                || blobs[0].Instance != 0) return false;
            IReadOnlyList<LegacyPptRecord> records;
            try {
                byte[] source = blobs[0].CopyRecordBytes();
                records = LegacyPptRecordReader.ReadSequence(source, 8,
                    blobs[0].PayloadLength,
                    new LegacyPptImportOptions());
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                return false;
            }
            if (records.Count(record =>
                    record.Type == RecordBlipCollection9ForWrite) > 1) {
                return false;
            }
            var rewritten = new List<byte[]>(records.Count
                + (collection == null ? 0 : 1));
            if (collection != null) rewritten.Add(collection);
            foreach (LegacyPptRecord record in records) {
                if (replaceExisting
                    && record.Type == RecordBlipCollection9ForWrite) {
                    continue;
                }
                rewritten.Add(record.CopyRecordBytes());
            }
            byte[] blob = BuildRecord(blobs[0].Version,
                blobs[0].Instance, blobs[0].Type, Concat(rewritten));
            var children = new List<byte[]>(binaryTag.Children.Count);
            foreach (LegacyPptRecord child in binaryTag.Children) {
                children.Add(ReferenceEquals(child, blobs[0])
                    ? blob
                    : child.CopyRecordBytes());
            }
            bytes = BuildRecord(binaryTag.Version, binaryTag.Instance,
                binaryTag.Type, Concat(children));
            return true;
        }

        private static byte[] BuildPictureBulletPpt9BinaryTag(
            byte[] collection) {
            byte[] name = BuildRecord(version: 0, instance: 0,
                RecordCString,
                System.Text.Encoding.Unicode.GetBytes(Ppt9TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob, collection);
            return BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { name, data });
        }

        internal static bool TryReadPictureBulletCatalog(
            PowerPointPresentation presentation,
            out LegacyPptWriterPictureBulletCatalog catalog,
            out string? reason) {
            if (presentation == null) throw new ArgumentNullException(
                nameof(presentation));
            PresentationPart? presentationPart = presentation
                .OpenXmlDocument.PresentationPart;
            if (presentationPart == null) {
                catalog = LegacyPptWriterPictureBulletCatalog.Empty;
                reason = "The Open XML presentation has no presentation part.";
                return false;
            }
            return TryReadPictureBulletCatalog(presentationPart,
                out catalog, out reason);
        }

        internal static bool TryReadPictureBulletCatalog(
            PresentationPart presentationPart,
            out LegacyPptWriterPictureBulletCatalog catalog,
            out string? reason) {
            if (presentationPart == null) throw new ArgumentNullException(
                nameof(presentationPart));
            var entries = new List<LegacyPptWriterPictureBulletEntry>();
            var byElement = new Dictionary<A.PictureBullet,
                LegacyPptWriterPictureBulletEntry>();
            var byHash = new Dictionary<string,
                List<LegacyPptWriterPictureBulletEntry>>(
                StringComparer.Ordinal);
            foreach (OpenXmlPart ownerPart in EnumeratePictureBulletParts(
                         presentationPart)) {
                OpenXmlPartRootElement? root = ownerPart.RootElement;
                if (root == null) continue;
                foreach (A.PictureBullet bullet in root
                             .Descendants<A.PictureBullet>()) {
                    if (!TryReadPictureBullet(ownerPart, bullet,
                            entries, byHash,
                            out LegacyPptWriterPictureBulletEntry? entry,
                            out reason) || entry == null) {
                        catalog = new LegacyPptWriterPictureBulletCatalog(
                            entries, byElement);
                        return false;
                    }
                    byElement.Add(bullet, entry);
                }
            }
            catalog = new LegacyPptWriterPictureBulletCatalog(entries,
                byElement);
            reason = null;
            return true;
        }

        private static IEnumerable<OpenXmlPart>
            EnumeratePictureBulletParts(OpenXmlPart root) {
            var pending = new Stack<OpenXmlPart>();
            var visited = new HashSet<OpenXmlPart>();
            pending.Push(root);
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part)) continue;
                yield return part;
                foreach (IdPartPair child in part.Parts) {
                    pending.Push(child.OpenXmlPart);
                }
            }
        }

        private static bool TryReadPictureBullet(OpenXmlPart ownerPart,
            A.PictureBullet bullet,
            ICollection<LegacyPptWriterPictureBulletEntry> entries,
            IDictionary<string, List<LegacyPptWriterPictureBulletEntry>>
                byHash,
            out LegacyPptWriterPictureBulletEntry? entry,
            out string? reason) {
            entry = null;
            reason = null;
            if (bullet.HasAttributes
                || bullet.ChildElements.Count != 1
                || bullet.GetFirstChild<A.Blip>() is not { } blip
                || !HasOnlyAttributes(blip, "embed")
                || blip.ChildElements.Count != 0
                || string.IsNullOrWhiteSpace(blip.Embed?.Value)) {
                reason = "A picture bullet must contain exactly one unmodified embedded DrawingML blip.";
                return false;
            }
            string relationshipId = blip.Embed?.Value ?? string.Empty;
            OpenXmlPart? relatedPart = ownerPart.Parts
                .Where(pair => string.Equals(pair.RelationshipId,
                    relationshipId, StringComparison.Ordinal))
                .Select(pair => pair.OpenXmlPart).FirstOrDefault();
            if (relatedPart is not ImagePart imagePart) {
                reason = "A picture bullet references a missing or non-image package part.";
                return false;
            }
            if (!TryGetPictureBulletBlipType(imagePart.ContentType,
                    out byte preferredType)) {
                reason = $"Picture bullets do not support image content type '{imagePart.ContentType}' in classic binary PowerPoint.";
                return false;
            }
            byte[] imageBytes;
            using (Stream stream = imagePart.GetStream(FileMode.Open,
                       FileAccess.Read)) {
                if (stream.Length <= 0
                    || stream.Length > MaximumPictureBulletBytes) {
                    reason = "A picture-bullet image is empty or exceeds the 64 MiB bounded binary-writing limit.";
                    return false;
                }
                using var memory = new MemoryStream(
                    checked((int)stream.Length));
                stream.CopyTo(memory);
                imageBytes = memory.ToArray();
            }
            byte[] blipRecord;
            try {
                blipRecord = OfficeArtBlipStoreEntryWriter
                    .CreateBlipRecord(imageBytes, imagePart.ContentType);
            } catch (Exception exception) when (exception
                is NotSupportedException or ArgumentException
                    or InvalidDataException or OverflowException) {
                reason = exception.Message;
                return false;
            }
            string key = imagePart.ContentType.ToLowerInvariant() + ":"
                + ComputePictureBulletHash(imageBytes);
            if (byHash.TryGetValue(key,
                    out List<LegacyPptWriterPictureBulletEntry>? matches)) {
                entry = matches.FirstOrDefault(candidate => candidate
                    .ImageBytes.SequenceEqual(imageBytes));
            }
            if (entry != null) return true;
            if (entries.Count >= 0x81) {
                reason = "Classic binary PowerPoint supports at most 129 distinct PPT9 picture bullets.";
                return false;
            }
            entry = new LegacyPptWriterPictureBulletEntry(
                checked((ushort)entries.Count), preferredType,
                imagePart.ContentType, imageBytes, blipRecord);
            entries.Add(entry);
            if (matches == null) {
                matches = new List<LegacyPptWriterPictureBulletEntry>();
                byHash.Add(key, matches);
            }
            matches.Add(entry);
            return true;
        }

        private static bool TryGetPictureBulletBlipType(
            string contentType, out byte value) {
            string normalized = contentType.Trim().ToLowerInvariant();
            if (normalized is "image/png" or "image/x-png") value = 0x06;
            else if (normalized is "image/jpeg" or "image/jpg") value = 0x05;
            else if (normalized is "image/x-emf" or "image/emf") value = 0x02;
            else if (normalized is "image/x-wmf" or "image/wmf") value = 0x03;
            else {
                value = 0;
                return false;
            }
            return true;
        }

        private static string ComputePictureBulletHash(byte[] bytes) {
            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(bytes));
        }

        internal sealed class LegacyPptWriterPictureBulletCatalog {
            private readonly IReadOnlyDictionary<A.PictureBullet,
                LegacyPptWriterPictureBulletEntry> _byElement;

            internal LegacyPptWriterPictureBulletCatalog(
                IEnumerable<LegacyPptWriterPictureBulletEntry> entries,
                IDictionary<A.PictureBullet,
                    LegacyPptWriterPictureBulletEntry> byElement) {
                Entries = new ReadOnlyCollection<
                    LegacyPptWriterPictureBulletEntry>(entries.ToArray());
                _byElement = new ReadOnlyDictionary<A.PictureBullet,
                    LegacyPptWriterPictureBulletEntry>(
                    new Dictionary<A.PictureBullet,
                        LegacyPptWriterPictureBulletEntry>(byElement));
            }

            internal static LegacyPptWriterPictureBulletCatalog Empty {
                get;
            } = new(Array.Empty<LegacyPptWriterPictureBulletEntry>(),
                new Dictionary<A.PictureBullet,
                    LegacyPptWriterPictureBulletEntry>());

            internal IReadOnlyList<LegacyPptWriterPictureBulletEntry>
                Entries { get; }

            internal bool TryGetIndex(A.PictureBullet bullet,
                out ushort index) {
                if (_byElement.TryGetValue(bullet,
                        out LegacyPptWriterPictureBulletEntry? entry)) {
                    index = entry.Index;
                    return true;
                }
                index = ushort.MaxValue;
                return false;
            }

            internal byte[]? BuildCollectionRecord() => Entries.Count == 0
                ? null
                : BuildContainer(RecordBlipCollection9ForWrite, instance: 0,
                    Entries.Select(entry => entry.BuildEntityRecord()));
        }

        internal sealed class LegacyPptWriterPictureBulletEntry {
            private readonly byte[] _imageBytes;
            private readonly byte[] _blipRecord;

            internal LegacyPptWriterPictureBulletEntry(ushort index,
                byte preferredType, string contentType, byte[] imageBytes,
                byte[] blipRecord) {
                Index = index;
                PreferredType = preferredType;
                ContentType = contentType;
                _imageBytes = imageBytes.ToArray();
                _blipRecord = blipRecord.ToArray();
            }

            internal ushort Index { get; }
            internal byte PreferredType { get; }
            internal string ContentType { get; }
            internal byte[] ImageBytes => _imageBytes.ToArray();

            internal byte[] BuildEntityRecord() {
                var payload = new byte[checked(2 + _blipRecord.Length)];
                payload[0] = PreferredType;
                Buffer.BlockCopy(_blipRecord, 0, payload, 2,
                    _blipRecord.Length);
                return BuildRecord(version: 0, instance: Index,
                    RecordBlipEntity9AtomForWrite, payload);
            }
        }
    }
}
