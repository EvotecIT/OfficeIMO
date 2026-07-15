using System.Text;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordDocumentInfoList = 0x07D0;
        private const ushort RecordEndDocumentAtom = 0x03EA;
        private const ushort RecordExternalHyperlink9 = 0x0FE4;
        private const ushort RecordExternalHyperlinkFlagsAtom = 0x1018;
        private const string Ppt9TagName = "___PPT9";

        internal static bool TryRewriteDocumentHyperlinkExtensions(
            LegacyPptRecord document,
            IReadOnlyList<LegacyPptWriterHyperlink> hyperlinks,
            bool replaceExisting, out byte[] bytes) {
            bytes = document.CopyRecordBytes();
            if (document.Version != 0x0F) return false;
            LegacyPptWriterHyperlink[] extensions = hyperlinks.Where(link =>
                link.ScreenTip != null || link.ExtensionFlags != 0).ToArray();
            LegacyPptRecord[] infoLists = document.Children.Where(record =>
                record.Type == RecordDocumentInfoList).ToArray();
            if (infoLists.Length > 1) return false;

            byte[]? rewrittenInfoList = null;
            if (infoLists.Length == 1) {
                if (!TryRewriteDocumentInfoList(infoLists[0], extensions,
                        replaceExisting, out rewrittenInfoList)) return false;
            } else if (extensions.Length > 0) {
                rewrittenInfoList = BuildContainer(RecordDocumentInfoList, instance: 0,
                    new[] { BuildContainer(RecordProgTags, instance: 0,
                        new[] { BuildPpt9BinaryTagRecord(extensions) }) });
            }

            var children = new List<byte[]>(document.Children.Count
                + (rewrittenInfoList != null && infoLists.Length == 0 ? 1 : 0));
            bool inserted = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (infoLists.Length == 1 && ReferenceEquals(child, infoLists[0])) {
                    children.Add(rewrittenInfoList!);
                    inserted = true;
                    continue;
                }
                if (rewrittenInfoList != null && infoLists.Length == 0 && !inserted
                    && (child.Type == RecordHeadersFooters
                        || (child.Type == RecordSlideListWithText && child.Instance != 1)
                        || child.Type == RecordEndDocumentAtom)) {
                    children.Add(rewrittenInfoList);
                    inserted = true;
                }
                children.Add(child.CopyRecordBytes());
            }
            if (rewrittenInfoList != null && !inserted) children.Add(rewrittenInfoList);
            bytes = BuildRecord(document.Version, document.Instance, document.Type,
                Concat(children));
            return true;
        }

        private static bool TryRewriteDocumentInfoList(LegacyPptRecord infoList,
            IReadOnlyList<LegacyPptWriterHyperlink> extensions,
            bool replaceExisting, out byte[] bytes) {
            bytes = infoList.CopyRecordBytes();
            if (infoList.Version != 0x0F || infoList.Instance != 0) return false;
            LegacyPptRecord[] tagLists = infoList.Children.Where(record =>
                record.Type == RecordProgTags).ToArray();
            if (tagLists.Length > 1) return false;
            byte[]? rewrittenTags = null;
            if (tagLists.Length == 1) {
                if (!TryRewriteDocumentProgTags(tagLists[0], extensions,
                        replaceExisting, out rewrittenTags)) return false;
            } else if (extensions.Count > 0) {
                rewrittenTags = BuildContainer(RecordProgTags, instance: 0,
                    new[] { BuildPpt9BinaryTagRecord(extensions) });
            }

            var children = new List<byte[]>(infoList.Children.Count
                + (rewrittenTags != null && tagLists.Length == 0 ? 1 : 0));
            foreach (LegacyPptRecord child in infoList.Children) {
                children.Add(tagLists.Length == 1 && ReferenceEquals(child, tagLists[0])
                    ? rewrittenTags!
                    : child.CopyRecordBytes());
            }
            if (rewrittenTags != null && tagLists.Length == 0) children.Add(rewrittenTags);
            bytes = BuildRecord(infoList.Version, infoList.Instance, infoList.Type,
                Concat(children));
            return true;
        }

        private static bool TryRewriteDocumentProgTags(LegacyPptRecord progTags,
            IReadOnlyList<LegacyPptWriterHyperlink> extensions,
            bool replaceExisting, out byte[] bytes) {
            bytes = progTags.CopyRecordBytes();
            if (progTags.Version != 0x0F || progTags.Instance != 0) return false;
            LegacyPptRecord[] ppt9Tags = progTags.Children.Where(IsPpt9BinaryTag).ToArray();
            if (ppt9Tags.Length > 1) return false;
            byte[]? rewrittenPpt9 = null;
            if (ppt9Tags.Length == 1) {
                if (!TryRewritePpt9BinaryTag(ppt9Tags[0], extensions,
                        replaceExisting, out rewrittenPpt9)) return false;
            } else if (extensions.Count > 0) {
                rewrittenPpt9 = BuildPpt9BinaryTagRecord(extensions);
            }

            var children = new List<byte[]>(progTags.Children.Count
                + (rewrittenPpt9 != null && ppt9Tags.Length == 0 ? 1 : 0));
            foreach (LegacyPptRecord child in progTags.Children) {
                children.Add(ppt9Tags.Length == 1 && ReferenceEquals(child, ppt9Tags[0])
                    ? rewrittenPpt9!
                    : child.CopyRecordBytes());
            }
            if (rewrittenPpt9 != null && ppt9Tags.Length == 0) children.Add(rewrittenPpt9);
            bytes = BuildRecord(progTags.Version, progTags.Instance, progTags.Type,
                Concat(children));
            return true;
        }

        private static bool TryRewritePpt9BinaryTag(LegacyPptRecord binaryTag,
            IReadOnlyList<LegacyPptWriterHyperlink> extensions,
            bool replaceExisting, out byte[] bytes) {
            bytes = binaryTag.CopyRecordBytes();
            LegacyPptRecord[] dataBlobs = binaryTag.Children.Where(record =>
                record.Type == RecordBinaryTagDataBlob).ToArray();
            if (binaryTag.Version != 0x0F || binaryTag.Instance != 0
                || dataBlobs.Length != 1 || dataBlobs[0].Version != 0
                || dataBlobs[0].Instance != 0) return false;
            IReadOnlyList<LegacyPptRecord> dataRecords;
            try {
                byte[] source = dataBlobs[0].CopyRecordBytes();
                dataRecords = LegacyPptRecordReader.ReadSequence(source, 8,
                    dataBlobs[0].PayloadLength, new LegacyPptImportOptions());
            } catch (InvalidDataException) {
                return false;
            }

            var rewrittenData = new List<byte[]>(dataRecords.Count + extensions.Count);
            int insertionIndex = -1;
            foreach (LegacyPptRecord record in dataRecords) {
                if (record.Type == RecordExternalHyperlink9) {
                    if (replaceExisting) continue;
                    rewrittenData.Add(record.CopyRecordBytes());
                    insertionIndex = rewrittenData.Count;
                    continue;
                }
                if (insertionIndex < 0 && IsPpt9PostHyperlinkRecord(record.Type)) {
                    insertionIndex = rewrittenData.Count;
                }
                rewrittenData.Add(record.CopyRecordBytes());
            }
            if (insertionIndex < 0) insertionIndex = rewrittenData.Count;
            rewrittenData.InsertRange(insertionIndex,
                extensions.Select(BuildExternalHyperlink9Record));
            byte[] dataBlob = BuildRecord(dataBlobs[0].Version, dataBlobs[0].Instance,
                dataBlobs[0].Type, Concat(rewrittenData));
            var children = new List<byte[]>(binaryTag.Children.Count);
            foreach (LegacyPptRecord child in binaryTag.Children) {
                children.Add(ReferenceEquals(child, dataBlobs[0])
                    ? dataBlob
                    : child.CopyRecordBytes());
            }
            bytes = BuildRecord(binaryTag.Version, binaryTag.Instance, binaryTag.Type,
                Concat(children));
            return true;
        }

        private static bool IsPpt9PostHyperlinkRecord(ushort type) =>
            type == 0x177A // PresentationAdvisorFlags9Atom
            || type == 0x1785 // EnvelopeData9Atom
            || type == 0x1784 // EnvelopeFlags9Atom
            || type == 0x177B // HtmlDocInfo9Atom
            || type == 0x177C // HtmlPublishInfoAtom
            || type == 0x177D // HtmlPublishInfo9Container
            || type == 0x177E // BroadcastDocInfo9Container
            || type == 0x0FAE; // OutlineTextProps9Container

        private static bool IsPpt9BinaryTag(LegacyPptRecord record) {
            if (record.Version != 0x0F || record.Instance != 0
                || record.Type != RecordProgBinaryTag) return false;
            LegacyPptRecord[] names = record.Children.Where(child =>
                child.Type == RecordCString && child.Instance == 0).ToArray();
            if (names.Length != 1 || (names[0].PayloadLength & 1) != 0) return false;
            try {
                return string.Equals(names[0].ReadUtf16Text().TrimEnd('\0'),
                    Ppt9TagName, StringComparison.Ordinal);
            } catch (InvalidDataException) {
                return false;
            }
        }

        private static byte[] BuildPpt9BinaryTagRecord(
            IReadOnlyList<LegacyPptWriterHyperlink> extensions) {
            byte[] tagName = BuildRecord(version: 0, instance: 0, RecordCString,
                Encoding.Unicode.GetBytes(Ppt9TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob,
                Concat(extensions.Select(BuildExternalHyperlink9Record)));
            return BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { tagName, data });
        }

        private static byte[] BuildExternalHyperlink9Record(
            LegacyPptWriterHyperlink hyperlink) {
            var referencePayload = new byte[4];
            WriteUInt32(referencePayload, 0, hyperlink.Id);
            var flagsPayload = new byte[4];
            WriteUInt32(flagsPayload, 0, hyperlink.ExtensionFlags);
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0,
                    RecordExternalHyperlinkAtom, referencePayload)
            };
            if (hyperlink.ScreenTip != null) {
                children.Add(BuildRecord(version: 0, instance: 0, RecordCString,
                    Encoding.Unicode.GetBytes(hyperlink.ScreenTip)));
            }
            children.Add(BuildRecord(version: 0, instance: 0,
                RecordExternalHyperlinkFlagsAtom, flagsPayload));
            return BuildContainer(RecordExternalHyperlink9, instance: 0, children);
        }
    }
}
