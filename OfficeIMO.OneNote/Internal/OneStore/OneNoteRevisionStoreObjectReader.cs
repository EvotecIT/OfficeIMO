namespace OfficeIMO.OneNote;

internal sealed class OneNoteRevisionStoreObjectReadResult {
    public List<OneNoteRevisionManifest> Revisions { get; } = new List<OneNoteRevisionManifest>();
    public List<OneNoteRevisionStoreObject> Objects { get; } = new List<OneNoteRevisionStoreObject>();
    public List<OneNoteFileDataStoreObject> FileDataObjects { get; } = new List<OneNoteFileDataStoreObject>();
}

internal static class OneNoteRevisionStoreObjectReader {
    private static readonly Guid FileDataHeader = new Guid("BDE316E7-2665-4511-A4C4-8D4D0B7A9EAC");
    private static readonly Guid FileDataFooter = new Guid("71FBA722-0F79-4A0B-BB13-899256426B24");

    public static OneNoteRevisionStoreObjectReadResult Read(
        Stream stream,
        OneNoteFileNodeList root,
        ulong declaredFileLength,
        OneNoteReaderOptions options) {
        var state = new ReaderState(stream, declaredFileLength, options);
        state.ProcessList(root, null);
        return state.Result;
    }

    private sealed class ReaderState {
        private readonly Stream _stream;
        private readonly ulong _declaredFileLength;
        private readonly OneNoteReaderOptions _options;
        private readonly HashSet<OneNoteFileNodeList> _visited = new HashSet<OneNoteFileNodeList>();
        private readonly Dictionary<OneNoteExtendedGuid, OneNoteRevisionManifest> _knownRevisions = new Dictionary<OneNoteExtendedGuid, OneNoteRevisionManifest>();
        private readonly Dictionary<ExtendedGuidKey, OneNoteJcid> _knownJcids = new Dictionary<ExtendedGuidKey, OneNoteJcid>();
        private int _nextRoleAssociationOrder;
        private long _totalAssetBytes;

        public ReaderState(Stream stream, ulong declaredFileLength, OneNoteReaderOptions options) {
            _stream = stream;
            _declaredFileLength = declaredFileLength;
            _options = options;
        }

        public OneNoteRevisionStoreObjectReadResult Result { get; } = new OneNoteRevisionStoreObjectReadResult();

        public void ProcessList(OneNoteFileNodeList list, OneNoteRevisionManifest? inheritedRevision) {
            if (!_visited.Add(list)) return;
            var frames = new Stack<ListFrame>();
            frames.Push(new ListFrame(list, inheritedRevision));

            while (frames.Count > 0) {
                ListFrame frame = frames.Peek();
                if (frame.NextNodeIndex >= frame.List.Nodes.Count) {
                    frames.Pop();
                    continue;
                }
                OneNoteFileNode node = frame.List.Nodes[frame.NextNodeIndex++];
                switch (node.RawId) {
                    case (ushort)OneNoteFileNodeId.RevisionManifestListStart:
                        frame.CurrentObjectSpace = ReadExtendedGuid(node.EncodedData.ToArray(24), 0);
                        break;
                    case (ushort)OneNoteFileNodeId.RevisionManifestStart4:
                    case (ushort)OneNoteFileNodeId.RevisionManifestStart6:
                    case (ushort)OneNoteFileNodeId.RevisionManifestStart7:
                        frame.CurrentRevision = ReadRevisionManifest(node);
                        frame.CurrentRevision.ObjectSpaceId = frame.CurrentObjectSpace;
                        frame.CurrentRevision.AddRoleAssociation(frame.CurrentRevision.ContextId, frame.CurrentRevision.Role, _nextRoleAssociationOrder++);
                        if (_knownRevisions.ContainsKey(frame.CurrentRevision.Id)) {
                            throw new OneNoteFormatException("ONENOTE_REVISION_ID", "A revision manifest identifier is duplicated.", node.FileOffset);
                        }
                        _knownRevisions.Add(frame.CurrentRevision.Id, frame.CurrentRevision);
                        Result.Revisions.Add(frame.CurrentRevision);
                        break;
                    case (ushort)OneNoteFileNodeId.RevisionRoleDeclaration:
                        ReadRevisionRoleDeclaration(node, frame.CurrentObjectSpace, false);
                        break;
                    case (ushort)OneNoteFileNodeId.RevisionRoleAndContextDeclaration:
                        ReadRevisionRoleDeclaration(node, frame.CurrentObjectSpace, true);
                        break;
                    case (ushort)OneNoteFileNodeId.GlobalIdTableStart:
                    case (ushort)OneNoteFileNodeId.GlobalIdTableStart2:
                        frame.GlobalIds.Clear();
                        break;
                    case (ushort)OneNoteFileNodeId.GlobalIdTableEntry:
                        ReadGlobalIdEntry(node, frame.GlobalIds);
                        break;
                    case (ushort)OneNoteFileNodeId.RootObjectReference2:
                    case (ushort)OneNoteFileNodeId.RootObjectReference3:
                        if (frame.CurrentRevision != null) ReadRootReference(node, frame.GlobalIds, frame.CurrentRevision);
                        break;
                    case (ushort)OneNoteFileNodeId.ObjectDeclarationWithRefCount:
                    case (ushort)OneNoteFileNodeId.ObjectDeclarationWithRefCount2:
                    case (ushort)OneNoteFileNodeId.ObjectDeclaration2RefCount:
                    case (ushort)OneNoteFileNodeId.ObjectDeclaration2LargeRefCount:
                    case (ushort)OneNoteFileNodeId.ReadOnlyObjectDeclaration2RefCount:
                    case (ushort)OneNoteFileNodeId.ReadOnlyObjectDeclaration2LargeRefCount:
                    case (ushort)OneNoteFileNodeId.ObjectRevisionWithRefCount:
                    case (ushort)OneNoteFileNodeId.ObjectRevisionWithRefCount2:
                        ReadObject(node, frame.GlobalIds, frame.CurrentRevision);
                        break;
                    case (ushort)OneNoteFileNodeId.ObjectDeclarationFileData3RefCount:
                    case (ushort)OneNoteFileNodeId.ObjectDeclarationFileData3LargeRefCount:
                        ReadFileDataDeclaration(node, frame.GlobalIds, frame.CurrentRevision);
                        break;
                    case (ushort)OneNoteFileNodeId.FileDataStoreObjectReference:
                        ReadFileDataStoreObject(node);
                        break;
                }

                if (node.ReferencedFileNodeList != null && _visited.Add(node.ReferencedFileNodeList)) {
                    OneNoteRevisionManifest? childRevision = node.RawId == (ushort)OneNoteFileNodeId.ObjectGroupListReference
                        ? frame.CurrentRevision
                        : null;
                    frames.Push(new ListFrame(node.ReferencedFileNodeList, childRevision));
                }
            }
        }

        private sealed class ListFrame {
            internal ListFrame(OneNoteFileNodeList list, OneNoteRevisionManifest? inheritedRevision) {
                List = list;
                CurrentRevision = inheritedRevision;
                CurrentObjectSpace = inheritedRevision?.ObjectSpaceId;
            }

            internal OneNoteFileNodeList List { get; }
            internal Dictionary<uint, Guid> GlobalIds { get; } = new Dictionary<uint, Guid>();
            internal OneNoteRevisionManifest? CurrentRevision { get; set; }
            internal OneNoteExtendedGuid? CurrentObjectSpace { get; set; }
            internal int NextNodeIndex { get; set; }
        }

        private OneNoteRevisionManifest ReadRevisionManifest(OneNoteFileNode node) {
            byte[] data = node.EncodedData.ToArray(128);
            OneNoteExtendedGuid id = ReadExtendedGuid(data, 0);
            OneNoteExtendedGuid dependency = ReadExtendedGuid(data, 20);
            int roleOffset = node.RawId == (ushort)OneNoteFileNodeId.RevisionManifestStart4 ? 48 : 40;
            int encryptionOffset = roleOffset + 4;
            OneNoteBinary.EnsureRange(data, encryptionOffset, 2);
            var manifest = new OneNoteRevisionManifest(id) {
                DependencyId = IsEmpty(dependency) ? null : dependency,
                Role = OneNoteBinary.ReadUInt32(data, roleOffset),
                IsEncrypted = OneNoteBinary.ReadUInt16(data, encryptionOffset) != 0
            };
            if (node.RawId == (ushort)OneNoteFileNodeId.RevisionManifestStart7) {
                manifest.ContextId = ReadExtendedGuid(data, 46);
            }
            return manifest;
        }

        private void ReadRevisionRoleDeclaration(
            OneNoteFileNode node,
            OneNoteExtendedGuid? currentObjectSpace,
            bool includesContext) {
            byte[] data = node.EncodedData.ToArray(includesContext ? 44 : 24);
            OneNoteExtendedGuid revisionId = ReadExtendedGuid(data, 0);
            uint role = OneNoteBinary.ReadUInt32(data, 20);
            if (role > ushort.MaxValue) {
                throw new OneNoteFormatException("ONENOTE_REVISION_ROLE", "A revision-role label has nonzero reserved high bytes.", node.FileOffset + 20);
            }
            if (!_knownRevisions.TryGetValue(revisionId, out OneNoteRevisionManifest? revision) ||
                currentObjectSpace == null ||
                !currentObjectSpace.Equals(revision.ObjectSpaceId)) {
                throw new OneNoteFormatException("ONENOTE_REVISION_ROLE_TARGET", "A revision-role declaration does not reference a preceding revision in the current object space.", node.FileOffset);
            }
            OneNoteExtendedGuid? contextId = null;
            if (includesContext) {
                OneNoteExtendedGuid context = ReadExtendedGuid(data, 24);
                if (!IsEmpty(context)) contextId = context;
            }
            revision.AddRoleAssociation(contextId, role, _nextRoleAssociationOrder++);
        }

        private static void ReadGlobalIdEntry(OneNoteFileNode node, Dictionary<uint, Guid> globalIds) {
            byte[] data = node.EncodedData.ToArray(20);
            OneNoteBinary.EnsureRange(data, 0, 20);
            uint index = OneNoteBinary.ReadUInt32(data, 0);
            if (index >= 0xFFFFFFU || globalIds.ContainsKey(index)) {
                throw new OneNoteFormatException("ONENOTE_GLOBAL_ID_INDEX", "A global-identification table index is invalid or duplicated.", node.FileOffset);
            }
            Guid identifier = OneNoteBinary.ReadGuid(data, 4);
            if (identifier == Guid.Empty) {
                throw new OneNoteFormatException("ONENOTE_GLOBAL_ID_GUID", "A global-identification table contains an empty GUID.", node.FileOffset + 4);
            }
            globalIds.Add(index, identifier);
        }

        private static void ReadRootReference(
            OneNoteFileNode node,
            Dictionary<uint, Guid> globalIds,
            OneNoteRevisionManifest revision) {
            byte[] data = node.EncodedData.ToArray(24);
            OneNoteExtendedGuid objectId;
            int roleOffset;
            if (node.RawId == (ushort)OneNoteFileNodeId.RootObjectReference3) {
                objectId = ReadExtendedGuid(data, 0);
                roleOffset = 20;
            } else {
                objectId = ResolveCompactId(data, 0, globalIds, node.FileOffset);
                roleOffset = 4;
            }
            uint role = OneNoteBinary.ReadUInt32(data, roleOffset);
            revision.RootObjects.Add(new OneNoteRootObjectReference(objectId, role));
        }

        private void ReadObject(
            OneNoteFileNode node,
            Dictionary<uint, Guid> globalIds,
            OneNoteRevisionManifest? revision) {
            if (Result.Objects.Count >= _options.MaxObjects) {
                throw new OneNoteFormatException("ONENOTE_OBJECT_LIMIT", "The object declaration limit was exceeded.", node.FileOffset);
            }
            if (node.ChunkReference == null || node.ChunkReference.IsNil || node.ChunkReference.IsZero) {
                throw new OneNoteFormatException("ONENOTE_OBJECT_REFERENCE", "An object declaration does not reference object data.", node.FileOffset);
            }

            byte[] data = node.EncodedData.ToArray(8192);
            int bodyOffset = node.ChunkReference.EncodedLength;
            OneNoteExtendedGuid id = ResolveCompactId(data, bodyOffset, globalIds, node.FileOffset + bodyOffset);
            OneNoteJcid jcid;
            uint referenceCount;
            bool isRevision = node.RawId == (ushort)OneNoteFileNodeId.ObjectRevisionWithRefCount ||
                              node.RawId == (ushort)OneNoteFileNodeId.ObjectRevisionWithRefCount2;

            if (node.RawId == (ushort)OneNoteFileNodeId.ObjectDeclarationWithRefCount ||
                node.RawId == (ushort)OneNoteFileNodeId.ObjectDeclarationWithRefCount2) {
                jcid = new OneNoteJcid(0x00020001U);
                int countOffset = bodyOffset + 10;
                referenceCount = node.RawId == (ushort)OneNoteFileNodeId.ObjectDeclarationWithRefCount
                    ? ReadByte(data, countOffset)
                    : OneNoteBinary.ReadUInt32(data, countOffset);
            } else if (isRevision) {
                if (!_knownJcids.TryGetValue(new ExtendedGuidKey(id), out jcid!)) jcid = new OneNoteJcid(0);
                int flagsOffset = bodyOffset + 4;
                if (node.RawId == (ushort)OneNoteFileNodeId.ObjectRevisionWithRefCount) {
                    byte flagsAndCount = ReadByte(data, flagsOffset);
                    referenceCount = (uint)(flagsAndCount >> 2);
                } else {
                    referenceCount = OneNoteBinary.ReadUInt32(data, flagsOffset + 4);
                }
            } else {
                jcid = new OneNoteJcid(OneNoteBinary.ReadUInt32(data, bodyOffset + 4));
                int countOffset = bodyOffset + 9;
                bool large = node.RawId == (ushort)OneNoteFileNodeId.ObjectDeclaration2LargeRefCount ||
                             node.RawId == (ushort)OneNoteFileNodeId.ReadOnlyObjectDeclaration2LargeRefCount;
                referenceCount = large ? OneNoteBinary.ReadUInt32(data, countOffset) : ReadByte(data, countOffset);
            }

            var record = new OneNoteRevisionStoreObject(id, jcid, node) {
                ReferenceCount = referenceCount,
                RevisionId = revision?.Id,
                IsRevision = isRevision
            };
            if (!revision?.IsEncrypted ?? true) {
                byte[] objectData = ReadReferencedBytes(node.ChunkReference, "object property set");
                record.RawPropertyData = OneNoteBinaryPayload.FromBytes(objectData);
                record.PropertySet = OneNotePropertySetReader.Read(objectData, globalIds, _options, node.ChunkReference.Offset);
            }
            Result.Objects.Add(record);
            if (jcid.Value != 0) _knownJcids[new ExtendedGuidKey(id)] = jcid;
        }

        private void ReadFileDataDeclaration(
            OneNoteFileNode node,
            Dictionary<uint, Guid> globalIds,
            OneNoteRevisionManifest? revision) {
            if (Result.Objects.Count >= _options.MaxObjects) {
                throw new OneNoteFormatException("ONENOTE_OBJECT_LIMIT", "The object declaration limit was exceeded.", node.FileOffset);
            }
            byte[] data = node.EncodedData.ToArray(8192);
            OneNoteExtendedGuid id = ResolveCompactId(data, 0, globalIds, node.FileOffset);
            var jcid = new OneNoteJcid(OneNoteBinary.ReadUInt32(data, 4));
            bool large = node.RawId == (ushort)OneNoteFileNodeId.ObjectDeclarationFileData3LargeRefCount;
            uint referenceCount = large ? OneNoteBinary.ReadUInt32(data, 8) : ReadByte(data, 8);
            int position = large ? 12 : 9;
            string reference = ReadStorageString(data, ref position, node.FileOffset);
            string extension = ReadStorageString(data, ref position, node.FileOffset);
            var record = new OneNoteRevisionStoreObject(id, jcid, node) {
                ReferenceCount = referenceCount,
                RevisionId = revision?.Id,
                FileDataReference = reference,
                FileExtension = extension
            };
            Result.Objects.Add(record);
            _knownJcids[new ExtendedGuidKey(id)] = jcid;
        }

        private void ReadFileDataStoreObject(OneNoteFileNode node) {
            if (node.ChunkReference == null || node.ChunkReference.IsNil || node.ChunkReference.IsZero) return;
            byte[] nodeData = node.EncodedData.ToArray(128);
            int guidOffset = node.ChunkReference.EncodedLength;
            Guid referenceId = OneNoteBinary.ReadGuid(nodeData, guidOffset);
            if (node.ChunkReference.Length < 52) {
                throw new OneNoteFormatException("ONENOTE_FILE_DATA_LENGTH", "A FileDataStoreObject is shorter than its required framing.", ToOffset(node.ChunkReference.Offset));
            }

            byte[] header = ReadReferencedRange(node.ChunkReference, 0, 36, "file-data store object header");
            ulong length = OneNoteBinary.ReadUInt64(header, 16);
            if (length > node.ChunkReference.Length - 52) {
                throw new OneNoteFormatException("ONENOTE_FILE_DATA_LENGTH", "A FileDataStoreObject payload length exceeds its containing frame.", ToOffset(node.ChunkReference.Offset + 16));
            }
            if (length > (ulong)_options.MaxAssetBytes || _totalAssetBytes > _options.MaxTotalAssetBytes - (long)length) {
                throw new OneNoteFormatException("ONENOTE_ASSET_LIMIT", "An embedded OneNote asset exceeds the configured materialization limits.", node.FileOffset);
            }
            if (length > int.MaxValue) {
                throw new OneNoteFormatException("ONENOTE_REFERENCED_DATA_SIZE", "A FileDataStoreObject payload is too large to materialize.", ToOffset(node.ChunkReference.Offset + 16));
            }

            byte[] footer = ReadReferencedRange(
                node.ChunkReference,
                node.ChunkReference.Length - 16,
                16,
                "file-data store object footer");
            if (OneNoteBinary.ReadGuid(header, 0) != FileDataHeader || OneNoteBinary.ReadGuid(footer, 0) != FileDataFooter) {
                throw new OneNoteFormatException("ONENOTE_FILE_DATA_FRAMING", "A FileDataStoreObject has invalid framing GUIDs.", ToOffset(node.ChunkReference.Offset));
            }

            byte[] payload = ReadReferencedRange(node.ChunkReference, 36, (int)length, "file-data store object payload");
            _totalAssetBytes += payload.Length;
            Result.FileDataObjects.Add(new OneNoteFileDataStoreObject(referenceId, OneNoteBinaryPayload.FromBytes(payload)));
        }

        private byte[] ReadReferencedBytes(OneNoteFileNodeChunkReference reference, string name) {
            if (reference.Length > int.MaxValue) {
                throw new OneNoteFormatException("ONENOTE_REFERENCED_DATA_SIZE", "The " + name + " is too large to materialize.", ToOffset(reference.Offset));
            }
            return ReadReferencedRange(reference, 0, (int)reference.Length, name);
        }

        private byte[] ReadReferencedRange(
            OneNoteFileNodeChunkReference reference,
            ulong relativeOffset,
            int length,
            string name) {
            if (reference.Offset > _declaredFileLength || reference.Length > _declaredFileLength - reference.Offset) {
                throw new OneNoteFormatException("ONENOTE_CHUNK_REFERENCE_BOUNDS", "The " + name + " lies outside the declared file length.", ToOffset(reference.Offset));
            }
            if (length < 0 || relativeOffset > reference.Length || (ulong)length > reference.Length - relativeOffset) {
                throw new OneNoteFormatException("ONENOTE_CHUNK_REFERENCE_BOUNDS", "The " + name + " lies outside its containing chunk reference.", ToOffset(reference.Offset));
            }
            ulong absoluteOffset = checked(reference.Offset + relativeOffset);
            _stream.Position = ToOffset(absoluteOffset);
            var data = new byte[length];
            int total = 0;
            while (total < data.Length) {
                int read = _stream.Read(data, total, data.Length - total);
                if (read <= 0) throw new OneNoteFormatException("ONENOTE_TRUNCATED_STRUCTURE", "The file ended while reading " + name + ".", ToOffset(absoluteOffset) + total);
                total += read;
            }
            return data;
        }
    }

    private static OneNoteExtendedGuid ReadExtendedGuid(byte[] data, int offset) {
        return new OneNoteExtendedGuid(OneNoteBinary.ReadGuid(data, offset), OneNoteBinary.ReadUInt32(data, offset + 16), 20);
    }

    private static OneNoteExtendedGuid ResolveCompactId(byte[] data, int offset, Dictionary<uint, Guid> globalIds, long absoluteOffset) {
        uint compact = OneNoteBinary.ReadUInt32(data, offset);
        byte value = (byte)(compact & 0xFFU);
        uint guidIndex = compact >> 8;
        if (!globalIds.TryGetValue(guidIndex, out Guid identifier)) {
            throw new OneNoteFormatException("ONENOTE_COMPACT_ID", "A CompactID references a missing global-identification table entry.", absoluteOffset);
        }
        return new OneNoteExtendedGuid(identifier, value, 4);
    }

    private static byte ReadByte(byte[] data, int offset) {
        OneNoteBinary.EnsureRange(data, offset, 1);
        return data[offset];
    }

    private static string ReadStorageString(byte[] data, ref int position, long absoluteOffset) {
        uint characterCount = OneNoteBinary.ReadUInt32(data, position);
        position += 4;
        if (characterCount > int.MaxValue / 2 || position > data.Length - (int)characterCount * 2) {
            throw new OneNoteFormatException("ONENOTE_STORAGE_STRING", "A StringInStorageBuffer length exceeds its containing structure.", absoluteOffset + position - 4);
        }
        string value = System.Text.Encoding.Unicode.GetString(data, position, (int)characterCount * 2);
        position += (int)characterCount * 2;
        return value;
    }

    private static bool IsEmpty(OneNoteExtendedGuid id) => id.Identifier == Guid.Empty && id.Value == 0;

    private static long ToOffset(ulong offset) {
        if (offset > long.MaxValue) throw new OneNoteFormatException("ONENOTE_OFFSET_RANGE", "A OneNote file offset exceeds the supported signed range.");
        return (long)offset;
    }

    private readonly struct ExtendedGuidKey : IEquatable<ExtendedGuidKey> {
        private readonly Guid _identifier;
        private readonly uint _value;

        public ExtendedGuidKey(OneNoteExtendedGuid id) {
            _identifier = id.Identifier;
            _value = id.Value;
        }

        public bool Equals(ExtendedGuidKey other) => _identifier == other._identifier && _value == other._value;
        public override bool Equals(object? obj) => obj is ExtendedGuidKey other && Equals(other);
        public override int GetHashCode() => (_identifier.GetHashCode() * 397) ^ _value.GetHashCode();
    }
}
