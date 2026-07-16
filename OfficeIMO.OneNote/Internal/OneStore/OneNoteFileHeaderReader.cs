namespace OfficeIMO.OneNote;

internal static class OneNoteFileHeaderReader {
    public static OneNoteFileHeader Read(byte[] prefix, long? actualFileLength, OneNoteReaderOptions options) {
        if (prefix == null) throw new ArgumentNullException(nameof(prefix));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (prefix.Length >= 4 && prefix[0] == (byte)'M' && prefix[1] == (byte)'S' && prefix[2] == (byte)'C' && prefix[3] == (byte)'F') {
            return new OneNoteFileHeader {
                FileKind = OneNoteFileKind.NotebookPackage,
                StorageFormat = OneNoteStorageFormat.NotebookPackage,
                ActualFileLength = actualFileLength
            };
        }
        OneNoteBinary.EnsureRange(prefix, 0, 64);

        Guid fileFormat = OneNoteBinary.ReadGuid(prefix, 48);
        if (fileFormat == OneNoteFormatConstants.RevisionStoreFormat) {
            return ReadRevisionStore(prefix, actualFileLength, options);
        }
        if (fileFormat == OneNoteFormatConstants.PackageStoreFormat) {
            return ReadPackageStore(prefix, actualFileLength, options);
        }

        throw new OneNoteFormatException(
            "ONENOTE_UNKNOWN_FILE_FORMAT",
            "The file does not contain a recognized MS-ONESTORE or package-store format identifier.",
            48);
    }

    private static OneNoteFileHeader ReadRevisionStore(byte[] data, long? actualFileLength, OneNoteReaderOptions options) {
        OneNoteBinary.EnsureRange(data, 0, OneNoteFormatConstants.RevisionStoreHeaderLength);

        Guid fileType = OneNoteBinary.ReadGuid(data, 0);
        OneNoteFileKind kind = ResolveDesktopFileKind(fileType);
        var header = new OneNoteFileHeader {
            FileKind = kind,
            StorageFormat = OneNoteStorageFormat.RevisionStore,
            FileTypeId = fileType,
            FileId = OneNoteBinary.ReadGuid(data, 16),
            LegacyFileVersionId = OneNoteBinary.ReadGuid(data, 32),
            FileFormatId = OneNoteBinary.ReadGuid(data, 48),
            ActualFileLength = actualFileLength,
            TransactionCount = OneNoteBinary.ReadUInt32(data, 96),
            AncestorId = OneNoteBinary.ReadGuid(data, 128),
            HashedChunkList = OneNoteBinary.ReadFileChunkReference64x32(data, 148),
            TransactionLog = OneNoteBinary.ReadFileChunkReference64x32(data, 160),
            RootFileNodeList = OneNoteBinary.ReadFileChunkReference64x32(data, 172),
            FreeChunkList = OneNoteBinary.ReadFileChunkReference64x32(data, 184),
            ExpectedFileLength = OneNoteBinary.ReadUInt64(data, 196),
            FileVersionId = OneNoteBinary.ReadGuid(data, 212),
            FileVersionGeneration = OneNoteBinary.ReadUInt64(data, 228),
            DenyReadFileVersionId = OneNoteBinary.ReadGuid(data, 236)
        };

        ValidateRevisionStoreHeader(data, header, options);
        return header;
    }

    private static OneNoteFileHeader ReadPackageStore(byte[] data, long? actualFileLength, OneNoteReaderOptions options) {
        OneNoteBinary.EnsureRange(data, 0, OneNoteFormatConstants.PackageStoreFixedPrefixLength + 1);
        Guid fileType = OneNoteBinary.ReadGuid(data, 0);
        if (fileType != OneNoteFormatConstants.SectionFileType) {
            throw new OneNoteFormatException(
                "ONENOTE_PACKAGE_FILE_TYPE",
                "The package-store header does not contain the required OneNote file-type identifier.",
                0);
        }

        uint streamHeader = OneNoteBinary.ReadUInt32(data, 68);
        uint headerType = streamHeader & 0x03;
        bool compound = (streamHeader & 0x04) != 0;
        uint streamObjectType = (streamHeader >> 3) & 0x3FFF;
        uint declaredLength = streamHeader >> 17;
        if (headerType != 0x02 || !compound || streamObjectType != 0x7A) {
            throw new OneNoteFormatException(
                "ONENOTE_PACKAGE_START",
                "The package-store header does not begin with the required packaging stream object.",
                68);
        }
        if (declaredLength == 0x7FFF) {
            throw new OneNoteFormatException(
                "ONENOTE_PACKAGE_LARGE_HEADER",
                "A package-store header with a large-length packaging prefix is not valid for the fixed OneNote envelope.",
                68);
        }

        OneNoteExtendedGuid storageIndex = OneNoteExtendedGuidReader.Read(data, 72);
        if (storageIndex.Identifier == Guid.Empty) {
            throw new OneNoteFormatException(
                "ONENOTE_PACKAGE_STORAGE_INDEX",
                "The package-store storage index identifier cannot be empty.",
                72);
        }

        int schemaOffset = 72 + storageIndex.EncodedLength;
        Guid cellSchema = OneNoteBinary.ReadGuid(data, schemaOffset);
        OneNoteFileKind kind = ResolveCellSchemaKind(cellSchema);
        int expectedPrefixPayloadLength = storageIndex.EncodedLength + 16;

        var header = new OneNoteFileHeader {
            FileKind = kind,
            StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage,
            FileTypeId = fileType,
            FileId = OneNoteBinary.ReadGuid(data, 16),
            LegacyFileVersionId = OneNoteBinary.ReadGuid(data, 32),
            FileFormatId = OneNoteBinary.ReadGuid(data, 48),
            ActualFileLength = actualFileLength,
            StorageIndexId = storageIndex,
            CellSchemaId = cellSchema
        };

        if (declaredLength != expectedPrefixPayloadLength) {
            ReportOrThrow(
                header,
                options,
                "ONENOTE_PACKAGE_PREFIX_LENGTH",
                "The packaging stream object length does not match the storage-index and cell-schema payload.",
                68);
        }

        return header;
    }

    private static void ValidateRevisionStoreHeader(byte[] data, OneNoteFileHeader header, OneNoteReaderOptions options) {
        uint expectedVersion = header.FileKind == OneNoteFileKind.Section ? 0x2AU : 0x1BU;
        for (int offset = 64; offset <= 76; offset += 4) {
            if (OneNoteBinary.ReadUInt32(data, offset) != expectedVersion) {
                ReportOrThrow(
                    header,
                    options,
                    "ONENOTE_FILE_VERSION",
                    "The revision-store header contains a file-format version that does not match its file type.",
                    offset);
            }
        }

        if (header.LegacyFileVersionId != Guid.Empty) {
            ReportOrThrow(
                header,
                options,
                "ONENOTE_LEGACY_VERSION_GUID",
                "The desktop revision-store legacy file-version identifier is not empty.",
                32);
        }

        if (!header.TransactionCount.HasValue || header.TransactionCount.Value == 0) {
            throw new OneNoteFormatException(
                "ONENOTE_TRANSACTION_COUNT",
                "The revision-store header declares no complete transactions.",
                96);
        }

        if (!header.ExpectedFileLength.HasValue || header.ExpectedFileLength.Value < OneNoteFormatConstants.RevisionStoreHeaderLength) {
            throw new OneNoteFormatException(
                "ONENOTE_EXPECTED_FILE_LENGTH",
                "The revision-store header declares an invalid expected file length.",
                196);
        }

        if (header.ActualFileLength.HasValue) {
            ulong actual = (ulong)header.ActualFileLength.Value;
            ulong expected = header.ExpectedFileLength.Value;
            if (actual < expected) {
                throw new OneNoteFormatException(
                    "ONENOTE_TRUNCATED_FILE",
                    "The source is shorter than the file length declared by its revision-store header.",
                    header.ActualFileLength.Value);
            }
            if (actual > expected) {
                header.AddDiagnostic(
                    "ONENOTE_TRAILING_DATA",
                    "The source contains trailing bytes beyond the file length declared by its revision-store header.",
                    (long)expected);
            }
        }

        ValidateRequiredReference(header.TransactionLog, "transaction log", 160, header.ExpectedFileLength.Value);
        ValidateRequiredReference(header.RootFileNodeList, "root file-node list", 172, header.ExpectedFileLength.Value);
        ValidateOptionalReference(header.HashedChunkList, "hashed chunk list", 148, header.ExpectedFileLength.Value);
        ValidateOptionalReference(header.FreeChunkList, "free chunk list", 184, header.ExpectedFileLength.Value);
    }

    private static OneNoteFileKind ResolveDesktopFileKind(Guid fileType) {
        if (fileType == OneNoteFormatConstants.SectionFileType) return OneNoteFileKind.Section;
        if (fileType == OneNoteFormatConstants.TableOfContentsFileType) return OneNoteFileKind.TableOfContents;
        throw new OneNoteFormatException(
            "ONENOTE_FILE_TYPE",
            "The revision-store header contains an unsupported OneNote file-type identifier.",
            0);
    }

    private static OneNoteFileKind ResolveCellSchemaKind(Guid schema) {
        if (schema == OneNoteFormatConstants.SectionCellSchema) return OneNoteFileKind.Section;
        if (schema == OneNoteFormatConstants.TableOfContentsCellSchema) return OneNoteFileKind.TableOfContents;
        throw new OneNoteFormatException(
            "ONENOTE_CELL_SCHEMA",
            "The package-store header contains an unsupported OneNote cell-schema identifier.");
    }

    private static void ValidateRequiredReference(OneNoteFileChunkReference? reference, string name, int offset, ulong expectedFileLength) {
        if (!reference.HasValue || reference.Value.IsNil || reference.Value.IsZero || reference.Value.Length == 0) {
            throw new OneNoteFormatException(
                "ONENOTE_REQUIRED_CHUNK_REFERENCE",
                "The revision-store " + name + " reference is missing.",
                offset);
        }
        ValidateReferenceBounds(reference.Value, name, offset, expectedFileLength);
    }

    private static void ValidateOptionalReference(OneNoteFileChunkReference? reference, string name, int offset, ulong expectedFileLength) {
        if (!reference.HasValue || reference.Value.IsNil || reference.Value.IsZero) return;
        ValidateReferenceBounds(reference.Value, name, offset, expectedFileLength);
    }

    private static void ValidateReferenceBounds(OneNoteFileChunkReference reference, string name, int offset, ulong expectedFileLength) {
        ulong length = reference.Length;
        if (reference.Offset > expectedFileLength || length > expectedFileLength - reference.Offset) {
            throw new OneNoteFormatException(
                "ONENOTE_CHUNK_REFERENCE_BOUNDS",
                "The revision-store " + name + " reference lies outside the declared file length.",
                offset);
        }
    }

    private static void ReportOrThrow(
        OneNoteFileHeader header,
        OneNoteReaderOptions options,
        string code,
        string message,
        long offset) {
        if (options.StrictHeaderValidation) {
            throw new OneNoteFormatException(code, message, offset);
        }
        header.AddDiagnostic(code, message, offset);
    }
}
