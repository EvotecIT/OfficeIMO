using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort OfficeArtDggContainerForPictures = 0xF000;
        private const ushort OfficeArtBStoreContainerForPictures = 0xF001;
        private const ushort OfficeArtFbseForPictures = 0xF007;

        private sealed class PreservingPictureStoreUpdate {
            private const int MaximumStoreEntryCount = 0x0FFF;
            private readonly LegacyPptPackage _package;
            private readonly Dictionary<int, int> _referenceCountDeltas = new();
            private bool _initialized;
            private int _existingEntryCount;
            private LegacyPptWriter.LegacyPptWriterPictureCatalog? _catalog;

            internal PreservingPictureStoreUpdate(LegacyPptPackage package) {
                _package = package ?? throw new ArgumentNullException(
                    nameof(package));
            }

            internal LegacyPptWriter.LegacyPptWriterPictureCatalog? Catalog =>
                _catalog;

            internal bool HasChanges => _referenceCountDeltas.Count > 0
                || _catalog?.Entries.Count > 0;

            internal bool TryPrepareBackground(LegacyPptRecord sourceOwner,
                LegacyPptWriter.LegacyPptWriterBackground background) {
                if (sourceOwner == null) throw new ArgumentNullException(
                    nameof(sourceOwner));
                if (background == null) throw new ArgumentNullException(
                    nameof(background));
                int? sourceStoreIndex = LegacyPptWriter
                    .ReadBackgroundPictureStoreIndex(sourceOwner);
                if (!sourceStoreIndex.HasValue
                    && !background.RequiresPictureCatalog) return true;
                if (!TryInitialize()) return false;

                if (sourceStoreIndex.HasValue) {
                    int index = sourceStoreIndex.Value;
                    if (index <= 0 || index > _existingEntryCount) return false;
                    _referenceCountDeltas.TryGetValue(index,
                        out int currentDelta);
                    _referenceCountDeltas[index] = checked(currentDelta - 1);
                }
                if (!background.RequiresPictureCatalog) return true;
                if (background.PictureFill == null
                    || background.PictureBytes.Length == 0
                    || string.IsNullOrWhiteSpace(
                        background.PictureContentType)) return false;
                return _catalog!.TryAdd(background.PictureFill,
                    background.PictureBytes, background.PictureContentType!,
                    out _);
            }

            internal bool TryBuildDocumentAndPictures(byte[]? currentDocumentBytes,
                out byte[] documentBytes, out byte[]? picturesBytes) {
                documentBytes = Array.Empty<byte>();
                picturesBytes = null;
                if (!HasChanges || !TryInitialize()) return false;
                LegacyPptRecord document;
                if (currentDocumentBytes != null) {
                    document = LegacyPptRecordReader.ReadSingle(
                        currentDocumentBytes, 0, new LegacyPptImportOptions());
                } else if (!TryReadDocument(_package,
                               out LegacyPptRecord? sourceDocument)
                           || sourceDocument == null) {
                    return false;
                } else {
                    document = sourceDocument;
                }

                LegacyPptRecord[] drawingGroups = document.Children.Where(
                    child => child.Type == RecordDrawingGroup).ToArray();
                if (drawingGroups.Length != 1
                    || !TryRewriteDrawingGroup(drawingGroups[0],
                        out byte[] rewrittenDrawingGroup)) return false;
                var documentChildren = new List<byte[]>(document.Children.Count);
                foreach (LegacyPptRecord child in document.Children) {
                    documentChildren.Add(ReferenceEquals(child, drawingGroups[0])
                        ? rewrittenDrawingGroup
                        : child.CopyRecordBytes());
                }
                documentBytes = BuildRecord(document.Version, document.Instance,
                    document.Type, Concat(documentChildren));

                if (_catalog!.Entries.Count == 0) return true;
                byte[] sourcePictures = _package.PicturesStream
                    ?? Array.Empty<byte>();
                byte[] appendedPictures = _catalog.BuildPicturesStream();
                picturesBytes = new byte[checked(sourcePictures.Length
                    + appendedPictures.Length)];
                Buffer.BlockCopy(sourcePictures, 0, picturesBytes, 0,
                    sourcePictures.Length);
                Buffer.BlockCopy(appendedPictures, 0, picturesBytes,
                    sourcePictures.Length, appendedPictures.Length);
                return true;
            }

            private bool TryInitialize() {
                if (_initialized) return _catalog != null;
                _initialized = true;
                if (!TryReadDocument(_package,
                        out LegacyPptRecord? document) || document == null) {
                    return false;
                }
                LegacyPptRecord[] drawingGroups = document.Children.Where(
                    child => child.Type == RecordDrawingGroup).ToArray();
                if (drawingGroups.Length != 1
                    || !TryReadStore(drawingGroups[0],
                        out LegacyPptRecord? store)) return false;
                _existingEntryCount = store?.Children.Count ?? 0;
                if (_existingEntryCount > MaximumStoreEntryCount) return false;
                uint delayedOffset = checked((uint)(
                    _package.PicturesStream?.Length ?? 0));
                _catalog = new LegacyPptWriter.LegacyPptWriterPictureCatalog(
                    _existingEntryCount, delayedOffset);
                return true;
            }

            private bool TryRewriteDrawingGroup(LegacyPptRecord drawingGroup,
                out byte[] bytes) {
                bytes = Array.Empty<byte>();
                LegacyPptRecord[] containers = drawingGroup.Children.Where(
                    child => child.Type == OfficeArtDggContainerForPictures)
                    .ToArray();
                if (containers.Length != 1
                    || !TryRewriteDggContainer(containers[0],
                        out byte[] rewrittenContainer)) return false;
                var children = new List<byte[]>(drawingGroup.Children.Count);
                foreach (LegacyPptRecord child in drawingGroup.Children) {
                    children.Add(ReferenceEquals(child, containers[0])
                        ? rewrittenContainer
                        : child.CopyRecordBytes());
                }
                bytes = BuildRecord(drawingGroup.Version, drawingGroup.Instance,
                    drawingGroup.Type, Concat(children));
                return true;
            }

            private bool TryRewriteDggContainer(LegacyPptRecord container,
                out byte[] bytes) {
                bytes = Array.Empty<byte>();
                LegacyPptRecord[] stores = container.Children.Where(child =>
                    child.Type == OfficeArtBStoreContainerForPictures).ToArray();
                if (stores.Length > 1) return false;
                LegacyPptRecord? store = stores.SingleOrDefault();
                if ((store?.Children.Count ?? 0) != _existingEntryCount) {
                    return false;
                }
                if (store == null && _referenceCountDeltas.Count > 0) {
                    return false;
                }

                byte[] rewrittenStore;
                if (!TryBuildStore(store, out rewrittenStore)) return false;
                var children = new List<byte[]>(container.Children.Count
                    + (store == null ? 1 : 0));
                bool inserted = false;
                foreach (LegacyPptRecord child in container.Children) {
                    if (store != null && ReferenceEquals(child, store)) {
                        children.Add(rewrittenStore);
                        inserted = true;
                    } else {
                        children.Add(child.CopyRecordBytes());
                        if (store == null && !inserted
                            && child.Type == OfficeArtDgg) {
                            children.Add(rewrittenStore);
                            inserted = true;
                        }
                    }
                }
                if (!inserted) children.Add(rewrittenStore);
                bytes = BuildRecord(container.Version, container.Instance,
                    container.Type, Concat(children));
                return true;
            }

            private bool TryBuildStore(LegacyPptRecord? source,
                out byte[] bytes) {
                bytes = Array.Empty<byte>();
                var entries = new List<byte[]>(checked(_existingEntryCount
                    + _catalog!.Entries.Count));
                if (source != null) {
                    if (source.Instance != _existingEntryCount
                        || source.Children.Any(child =>
                            child.Type != OfficeArtFbseForPictures
                            || child.PayloadLength < 36)) return false;
                    for (int index = 0; index < source.Children.Count; index++) {
                        LegacyPptRecord entry = source.Children[index];
                        byte[] entryBytes = entry.CopyRecordBytes();
                        if (_referenceCountDeltas.TryGetValue(index + 1,
                                out int delta)) {
                            uint sourceCount = ReadUInt32(entryBytes, 32);
                            long updated = sourceCount + (long)delta;
                            if (updated < 0 || updated > uint.MaxValue) {
                                return false;
                            }
                            WriteUInt32(entryBytes, 32,
                                checked((uint)updated));
                        }
                        entries.Add(entryBytes);
                    }
                }
                if (_referenceCountDeltas.Keys.Any(index =>
                        index <= 0 || index > _existingEntryCount)) return false;
                entries.AddRange(_catalog.BuildDelayedStoreEntries());
                if (entries.Count == 0
                    || entries.Count > MaximumStoreEntryCount) return false;
                bytes = BuildRecord(source?.Version ?? 0x0F,
                    checked((ushort)entries.Count),
                    OfficeArtBStoreContainerForPictures, Concat(entries));
                return true;
            }

            private static bool TryReadStore(LegacyPptRecord drawingGroup,
                out LegacyPptRecord? store) {
                store = null;
                LegacyPptRecord[] containers = drawingGroup.Children.Where(
                    child => child.Type == OfficeArtDggContainerForPictures)
                    .ToArray();
                if (containers.Length != 1) return false;
                LegacyPptRecord[] stores = containers[0].Children.Where(child =>
                    child.Type == OfficeArtBStoreContainerForPictures).ToArray();
                if (stores.Length > 1) return false;
                store = stores.SingleOrDefault();
                return store == null
                    || store.Instance == store.Children.Count
                    && store.Children.All(child =>
                        child.Type == OfficeArtFbseForPictures
                        && child.PayloadLength >= 36);
            }
        }
    }
}
