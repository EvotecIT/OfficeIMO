using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordSoundCollection = 0x07E4;
        private const ushort RecordSoundCollectionAtom = 0x07E5;
        private const ushort RecordSound = 0x07E6;
        private const ushort RecordDrawingGroup = 0x040B;

        private static bool TryAppendNewSounds(LegacyPptPackage package,
            byte[]? currentDocumentBytes,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterSound> sounds,
            out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (sounds.Count == 0) return false;
            LegacyPptRecord document;
            if (currentDocumentBytes != null) {
                document = LegacyPptRecordReader.ReadSingle(currentDocumentBytes,
                    0, new LegacyPptImportOptions());
            } else if (!TryReadDocument(package, out LegacyPptRecord? source)
                       || source == null) {
                return false;
            } else {
                document = source;
            }

            LegacyPptRecord[] collections = document.Children.Where(record =>
                record.Type == RecordSoundCollection).ToArray();
            if (collections.Length > 1) return false;
            uint greatestNewId = sounds.Max(sound => sound.Id);
            byte[] rewrittenCollection;
            if (collections.Length == 0) {
                var seedPayload = new byte[4];
                WriteInt32(seedPayload, 0, checked((int)greatestNewId));
                var children = new List<byte[]> {
                    BuildRecord(version: 0, instance: 0,
                        RecordSoundCollectionAtom, seedPayload)
                };
                children.AddRange(sounds.Select(
                    LegacyPptWriter.BuildSoundRecord));
                rewrittenCollection = BuildRecord(version: 0x0F, instance: 5,
                    RecordSoundCollection, Concat(children));
            } else {
                LegacyPptRecord collection = collections[0];
                LegacyPptRecord[] atoms = collection.Children.Where(record =>
                    record.Type == RecordSoundCollectionAtom).ToArray();
                if (collection.Version != 0x0F || collection.Instance != 5
                    || atoms.Length != 1 || atoms[0].Version != 0
                    || atoms[0].Instance != 0 || atoms[0].PayloadLength != 4
                    || atoms[0].ReadInt32(0) <= 0
                    || collection.Children.Any(child =>
                        child.Type != RecordSoundCollectionAtom
                        && child.Type != RecordSound)) {
                    return false;
                }
                var children = new List<byte[]>(collection.Children.Count
                    + sounds.Count);
                foreach (LegacyPptRecord child in collection.Children) {
                    if (ReferenceEquals(child, atoms[0])) {
                        byte[] atom = child.CopyRecordBytes();
                        WriteInt32(atom, 8, Math.Max(child.ReadInt32(0),
                            checked((int)greatestNewId)));
                        children.Add(atom);
                    } else {
                        children.Add(child.CopyRecordBytes());
                    }
                }
                children.AddRange(sounds.Select(
                    LegacyPptWriter.BuildSoundRecord));
                rewrittenCollection = BuildRecord(collection.Version,
                    collection.Instance, collection.Type, Concat(children));
            }

            var documentChildren = new List<byte[]>(document.Children.Count + 1);
            bool inserted = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (collections.Length == 1
                    && ReferenceEquals(child, collections[0])) {
                    documentChildren.Add(rewrittenCollection);
                    inserted = true;
                } else {
                    if (!inserted && collections.Length == 0
                        && child.Type == RecordDrawingGroup) {
                        documentChildren.Add(rewrittenCollection);
                        inserted = true;
                    }
                    documentChildren.Add(child.CopyRecordBytes());
                }
            }
            if (!inserted) documentChildren.Add(rewrittenCollection);
            bytes = BuildRecord(document.Version, document.Instance,
                document.Type, Concat(documentChildren));
            return true;
        }
    }
}
