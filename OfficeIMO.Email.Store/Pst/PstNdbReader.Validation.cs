namespace OfficeIMO.Email.Store;

internal sealed partial class PstNdbReader {
    internal EmailStoreStructuralValidationResult ValidateStructure(
        EmailStoreValidationOptions options,
        CancellationToken cancellationToken) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var context = new StructuralValidationContext(
            _stream, _header, options, cancellationToken);
        return context.Validate();
    }

    private sealed class StructuralValidationContext {
        private readonly Stream _stream;
        private readonly PstHeader _header;
        private readonly EmailStoreValidationOptions _options;
        private readonly CancellationToken _cancellationToken;
        private readonly List<EmailStoreDiagnostic> _diagnostics =
            new List<EmailStoreDiagnostic>();
        private readonly HashSet<long> _visitedPages = new HashSet<long>();
        private readonly List<PstBlockReference> _blocks = new List<PstBlockReference>();
        private int _pagesExamined;
        private int _blocksExamined;
        private long _bytesExamined;
        private int _failures;
        private bool _wasTruncated;

        internal StructuralValidationContext(Stream stream, PstHeader header,
            EmailStoreValidationOptions options, CancellationToken cancellationToken) {
            _stream = stream;
            _header = header;
            _options = options;
            _cancellationToken = cancellationToken;
        }

        internal EmailStoreStructuralValidationResult Validate() {
            TraversePage(_header.BbtRootOffset, _header.BbtRootBid, 0x80,
                isBlockTree: true, depth: 0);
            TraversePage(_header.NbtRootOffset, _header.NbtRootBid, 0x81,
                isBlockTree: false, depth: 0);
            foreach (PstBlockReference block in _blocks) {
                if (!TryReserveBlock(block)) break;
                ValidateBlock(block);
            }
            return new EmailStoreStructuralValidationResult(
                supported: true, _pagesExamined, _blocksExamined, _bytesExamined,
                _failures, _wasTruncated, _diagnostics.ToArray());
        }

        private void TraversePage(long offset, ulong expectedBid, byte expectedType,
            bool isBlockTree, int depth) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (depth > 64) {
                AddPageFailure("EMAIL_STORE_PST_PAGE_DEPTH",
                    "A PST B-tree exceeds the structural validation depth limit.", offset);
                return;
            }
            if (!_visitedPages.Add(offset)) {
                AddPageFailure("EMAIL_STORE_PST_PAGE_REUSED",
                    "A PST B-tree page is referenced more than once.", offset);
                return;
            }
            if (!TryReservePage()) return;

            byte[] page;
            try {
                if (offset < 0 || offset % _header.PageSize != 0 ||
                    offset > _stream.Length - _header.PageSize) {
                    throw new InvalidDataException(
                        "A PST B-tree page is misaligned or outside the source stream.");
                }
                page = PstBinary.ReadAt(_stream, offset, _header.PageSize);
            } catch (Exception exception) when (
                exception is InvalidDataException || exception is EndOfStreamException) {
                AddPageFailure("EMAIL_STORE_PST_PAGE_BOUNDS", exception.Message, offset);
                return;
            }

            bool invalid = !ValidatePageTrailer(page, offset, expectedBid, expectedType);
            int metadataOffset = _header.PageSize - _header.PageTrailerSize -
                _header.BTreeMetadataSize;
            int count;
            int entrySize;
            int level;
            try {
                if (_header.Variant == PstVariant.Unicode4K) {
                    count = PstBinary.UInt16(page, metadataOffset);
                    entrySize = page[metadataOffset + 4];
                    level = page[metadataOffset + 5];
                } else {
                    count = page[metadataOffset];
                    entrySize = page[metadataOffset + 2];
                    level = page[metadataOffset + 3];
                }
                if (entrySize <= 0 || checked(count * entrySize) > metadataOffset) {
                    throw new InvalidDataException(
                        "A PST B-tree page has an invalid entry layout.");
                }
            } catch (Exception exception) when (
                exception is InvalidDataException || exception is OverflowException) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_LAYOUT", exception.Message, offset);
                _failures++;
                return;
            }

            int keySize = _header.IsUnicode ? 8 : 4;
            int bidSize = _header.IsUnicode ? 8 : 4;
            int minimumEntrySize = level > 0
                ? keySize + bidSize + (_header.IsUnicode ? 8 : 4)
                : isBlockTree
                    ? (_header.IsUnicode ? 20 : 12)
                    : (_header.IsUnicode ? 28 : 16);
            if (entrySize < minimumEntrySize) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_ENTRY_TRUNCATED",
                    "A PST B-tree entry is shorter than its required fields.", offset);
                _failures++;
                return;
            }
            if (!ValidateKeyOrder(page, count, entrySize, keySize, isBlockTree, offset)) {
                invalid = true;
            }

            if (level > 0) {
                for (int index = 0; index < count; index++) {
                    int entryOffset = checked(index * entrySize);
                    ulong childBid = _header.IsUnicode
                        ? PstBinary.UInt64(page, entryOffset + keySize)
                        : PstBinary.UInt32(page, entryOffset + keySize);
                    long childOffset = _header.IsUnicode
                        ? checked((long)PstBinary.UInt64(page, entryOffset + keySize + bidSize))
                        : PstBinary.UInt32(page, entryOffset + keySize + bidSize);
                    TraversePage(childOffset, childBid, expectedType, isBlockTree, depth + 1);
                }
            } else if (isBlockTree) {
                for (int index = 0; index < count; index++) {
                    if (_blocks.Count >= _options.MaxStructuralBlocks) {
                        _wasTruncated = true;
                        break;
                    }
                    try {
                        _blocks.Add(ReadBlockReference(page, index * entrySize, entrySize));
                    } catch (Exception exception) when (
                        exception is InvalidDataException || exception is OverflowException) {
                        AddPageDiagnostic("EMAIL_STORE_PST_BBT_ENTRY_INVALID",
                            exception.Message, offset);
                        invalid = true;
                    }
                }
            }
            if (invalid) _failures++;
        }

        private bool ValidatePageTrailer(byte[] page, long offset,
            ulong expectedBid, byte expectedType) {
            int trailerOffset = _header.PageSize - _header.PageTrailerSize;
            bool valid = true;
            if (page[trailerOffset] != expectedType ||
                page[trailerOffset + 1] != expectedType) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_TYPE",
                    "A PST B-tree page has an unexpected or non-repeated page type.", offset);
                valid = false;
            }
            ushort actualSignature = PstBinary.UInt16(page, trailerOffset + 2);
            int crcOffset = _header.Variant == PstVariant.Ansi
                ? trailerOffset + 8
                : trailerOffset + 4;
            int bidOffset = _header.Variant == PstVariant.Ansi
                ? trailerOffset + 4
                : trailerOffset + 8;
            uint actualCrc = PstBinary.UInt32(page, crcOffset);
            ulong actualBid = _header.IsUnicode
                ? PstBinary.UInt64(page, bidOffset)
                : PstBinary.UInt32(page, bidOffset);
            uint expectedCrc = PstCrc32.Compute(page, trailerOffset);
            if (actualCrc != expectedCrc) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_CRC",
                    "A PST B-tree page CRC does not match its page data.", offset);
                valid = false;
            }
            if (actualBid != expectedBid) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_BID",
                    "A PST B-tree page trailer BID does not match its parent reference.", offset);
                valid = false;
            }
            if (actualSignature != PstSignature.Compute(offset, actualBid)) {
                AddPageDiagnostic("EMAIL_STORE_PST_PAGE_SIGNATURE",
                    "A PST B-tree page signature does not match its offset and BID.", offset);
                valid = false;
            }
            return valid;
        }

        private bool ValidateKeyOrder(byte[] page, int count, int entrySize,
            int keySize, bool isBlockTree, long offset) {
            ulong previous = 0;
            for (int index = 0; index < count; index++) {
                int entryOffset = index * entrySize;
                ulong key = keySize == 8
                    ? PstBinary.UInt64(page, entryOffset)
                    : PstBinary.UInt32(page, entryOffset);
                if (isBlockTree) key = PstBinary.NormalizeBid(key);
                if (index > 0 && key <= previous) {
                    AddPageDiagnostic("EMAIL_STORE_PST_PAGE_KEY_ORDER",
                        "PST B-tree page keys are not strictly increasing.", offset);
                    return false;
                }
                previous = key;
            }
            return true;
        }

        private PstBlockReference ReadBlockReference(byte[] page, int offset, int entrySize) {
            int minimum = _header.IsUnicode ? 20 : 12;
            if (entrySize < minimum) {
                throw new InvalidDataException("A PST BBT leaf entry is truncated.");
            }
            ulong bid = _header.IsUnicode
                ? PstBinary.UInt64(page, offset)
                : PstBinary.UInt32(page, offset);
            long blockOffset = _header.IsUnicode
                ? checked((long)PstBinary.UInt64(page, offset + 8))
                : PstBinary.UInt32(page, offset + 4);
            int storedLength = PstBinary.UInt16(
                page, offset + (_header.IsUnicode ? 16 : 8));
            int decodedLength = _header.Variant == PstVariant.Unicode4K
                ? PstBinary.UInt16(page, offset + 18)
                : storedLength;
            if (decodedLength <= 0) decodedLength = storedLength;
            return new PstBlockReference(bid, blockOffset, storedLength, decodedLength);
        }

        private void ValidateBlock(PstBlockReference block) {
            bool invalid = false;
            string location = string.Concat("block/0x",
                PstBinary.NormalizeBid(block.Bid).ToString("X", CultureInfo.InvariantCulture));
            try {
                if (block.DataLength < 0 || block.Offset < 0 ||
                    block.Offset % _header.BlockAlignment != 0) {
                    throw new InvalidDataException(
                        "A PST block has an invalid length, offset, or alignment.");
                }
                int allocationLength = PstBinary.Align(
                    checked(block.DataLength + _header.BlockTrailerSize),
                    _header.BlockAlignment);
                if (block.Offset > _stream.Length - allocationLength) {
                    throw new InvalidDataException(
                        "A PST block allocation extends outside the source stream.");
                }
                byte[] payload = PstBinary.ReadAt(_stream, block.Offset, block.DataLength);
                long trailerPosition = checked(block.Offset + allocationLength - _header.BlockTrailerSize);
                byte[] trailer = PstBinary.ReadAt(
                    _stream, trailerPosition, _header.BlockTrailerSize);
                int storedLength = PstBinary.UInt16(trailer, 0);
                ushort actualSignature = PstBinary.UInt16(trailer, 2);
                uint actualCrc = PstBinary.UInt32(trailer, 4);
                ulong actualBid = _header.IsUnicode
                    ? PstBinary.UInt64(trailer, 8)
                    : PstBinary.UInt32(trailer, 8);
                if (storedLength != block.DataLength) {
                    AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_LENGTH",
                        "A PST block trailer length does not match its BBT entry.", location);
                    invalid = true;
                }
                if (_header.Variant == PstVariant.Unicode4K &&
                    PstBinary.UInt16(trailer, 18) != block.DecodedLength) {
                    AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_DECODED_LENGTH",
                        "A 4K OST block trailer decoded length does not match its BBT entry.", location);
                    invalid = true;
                }
                if (PstBinary.NormalizeBid(actualBid) != PstBinary.NormalizeBid(block.Bid)) {
                    AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_BID",
                        "A PST block trailer BID does not match its BBT entry.", location);
                    invalid = true;
                }
                if (actualSignature != PstSignature.Compute(block.Offset, actualBid)) {
                    AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_SIGNATURE",
                        "A PST block signature does not match its offset and BID.", location);
                    invalid = true;
                }
                if (actualCrc != PstCrc32.Compute(payload)) {
                    AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_CRC",
                        "A PST block CRC does not match its stored payload.", location);
                    invalid = true;
                }
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is EndOfStreamException ||
                exception is OverflowException) {
                AddBlockDiagnostic("EMAIL_STORE_PST_BLOCK_BOUNDS",
                    exception.Message, location);
                invalid = true;
            }
            if (invalid) _failures++;
        }

        private bool TryReservePage() {
            if (_pagesExamined >= _options.MaxStructuralPages ||
                _bytesExamined > _options.MaxStructuralBytes - _header.PageSize) {
                _wasTruncated = true;
                return false;
            }
            _pagesExamined++;
            _bytesExamined += _header.PageSize;
            return true;
        }

        private bool TryReserveBlock(PstBlockReference block) {
            long bytes = checked((long)block.DataLength + _header.BlockTrailerSize);
            if (_blocksExamined >= _options.MaxStructuralBlocks ||
                bytes > _options.MaxStructuralBytes ||
                _bytesExamined > _options.MaxStructuralBytes - bytes) {
                _wasTruncated = true;
                return false;
            }
            _blocksExamined++;
            _bytesExamined += bytes;
            return true;
        }

        private void AddPageFailure(string code, string message, long offset) {
            AddPageDiagnostic(code, message, offset);
            _failures++;
        }

        private void AddPageDiagnostic(string code, string message, long offset) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                code, message, EmailStoreDiagnosticSeverity.Error,
                string.Concat("page/0x", offset.ToString("X", CultureInfo.InvariantCulture))));
        }

        private void AddBlockDiagnostic(string code, string message, string location) {
            _diagnostics.Add(new EmailStoreDiagnostic(
                code, message, EmailStoreDiagnosticSeverity.Error, location));
        }
    }
}
