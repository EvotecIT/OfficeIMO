using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocPictureReader {
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtFbse = 0xF007;
        private const short MmShape = 0x0064;
        private const short MmShapeFile = 0x0066;
        private const double PixelsPerTwip = 1D / 15D;

        internal static LegacyDocPictureReadResult Read(
            byte[] dataStream,
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            int supportedStoryCharacterCount,
            int maximumDecodedImageBytes) {
            if (maximumDecodedImageBytes <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maximumDecodedImageBytes));
            }
            if (dataStream.Length == 0 || supportedStoryCharacterCount <= 0) {
                return LegacyDocPictureReadResult.Empty;
            }

            var pictures = new Dictionary<int, LegacyDocPicture>();
            var picturesByDataOffset = new Dictionary<int, LegacyDocPicture>();
            var ranges = new Dictionary<int, int>();
            int decodedImageBytes = 0;
            string? warning = null;
            foreach (LegacyDocTextCharacter character in characters) {
                if (character.CharacterPosition < 0 || character.CharacterPosition >= supportedStoryCharacterCount
                    || character.Character != '\u0001') {
                    continue;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, character.FileOffset);
                if (format.PictureDataOffset == null) {
                    continue;
                }

                int offset = format.PictureDataOffset.Value;
                if (picturesByDataOffset.TryGetValue(offset, out LegacyDocPicture? cachedPicture)) {
                    pictures[character.CharacterPosition] = cachedPicture;
                    continue;
                }
                int remainingImageBytes = maximumDecodedImageBytes - decodedImageBytes;
                if (!TryReadPicture(dataStream, offset, remainingImageBytes,
                        out LegacyDocPicture? picture, out int consumedLength,
                        out string? pictureWarning)) {
                    warning ??= pictureWarning;
                    continue;
                }

                decodedImageBytes = checked(decodedImageBytes + picture!.ImageByteCount);
                picturesByDataOffset.Add(offset, picture);
                pictures[character.CharacterPosition] = picture!;
                ranges[offset] = Math.Max(ranges.TryGetValue(offset, out int previousLength) ? previousLength : 0, consumedLength);
            }

            bool fullyProjected = pictures.Count > 0 && CoversDataStream(ranges, dataStream.Length);
            return new LegacyDocPictureReadResult(pictures, fullyProjected, warning);
        }

        internal static bool TryCreatePictureRun(
            LegacyDocTextCharacter character,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition,
            LegacyDocCharacterFormat format,
            out LegacyDocTextRun? pictureRun) {
            pictureRun = null;
            if (character.Character != '\u0001'
                || !picturesByCharacterPosition.TryGetValue(character.CharacterPosition, out LegacyDocPicture? picture)) {
                return false;
            }

            pictureRun = new LegacyDocTextRun(
                string.Empty,
                bold: false,
                italic: false,
                strike: false,
                doubleStrike: false,
                outline: false,
                shadow: false,
                emboss: false,
                imprint: false,
                hidden: false,
                noProof: false,
                caps: null,
                verticalPosition: null,
                underline: null,
                highlight: null,
                fontSizeHalfPoints: null,
                colorHex: null,
                fontFamily: null,
                characterPositions: new[] { character.CharacterPosition },
                picture: picture,
                revision: format.Revision);
            return true;
        }

        private static bool TryReadPicture(
            byte[] data,
            int offset,
            int maximumDecodedImageBytes,
            out LegacyDocPicture? picture,
            out int consumedLength,
            out string? warning) {
            picture = null;
            consumedLength = 0;
            warning = null;
            if (offset < 0 || offset > data.Length - 68) {
                warning = $"A picture character points outside the DOC Data stream at offset {offset}.";
                return false;
            }

            int lcb = LegacyDocFib.ReadInt32(data, offset);
            ushort cbHeader = LegacyDocFib.ReadUInt16(data, offset + 4);
            short pictureFormat = unchecked((short)LegacyDocFib.ReadUInt16(data, offset + 6));
            if (lcb < 76 || lcb > data.Length - offset || cbHeader != 68
                || (pictureFormat != MmShape && pictureFormat != MmShapeFile)) {
                warning = $"The picture data at DOC Data stream offset {offset} has an unsupported or truncated PICF header.";
                return false;
            }

            short widthGoal = unchecked((short)LegacyDocFib.ReadUInt16(data, offset + 28));
            short heightGoal = unchecked((short)LegacyDocFib.ReadUInt16(data, offset + 30));
            ushort widthScale = LegacyDocFib.ReadUInt16(data, offset + 32);
            ushort heightScale = LegacyDocFib.ReadUInt16(data, offset + 34);
            if (widthGoal <= 0 || heightGoal <= 0 || widthScale == 0 || heightScale == 0) {
                warning = $"The picture data at DOC Data stream offset {offset} has invalid display dimensions.";
                return false;
            }

            int shapeOffset = offset + cbHeader;
            if (pictureFormat == MmShapeFile) {
                if (shapeOffset >= offset + lcb) {
                    warning = $"The shape-file picture at DOC Data stream offset {offset} is truncated before its name.";
                    return false;
                }

                shapeOffset = checked(shapeOffset + 1 + data[shapeOffset]);
            }

            int pictureEnd = checked(offset + lcb);
            if (!TryReadOfficeArtHeader(data, shapeOffset, pictureEnd, out ushort shapeType, out _, out int shapeLength)
                || shapeType != OfficeArtSpContainer) {
                warning = $"The picture data at DOC Data stream offset {offset} has no supported OfficeArt inline-shape container.";
                return false;
            }

            int blipStoreOffset = checked(shapeOffset + 8 + shapeLength);
            if (!TryReadOfficeArtHeader(data, blipStoreOffset, pictureEnd, out ushort recordType, out ushort recordInstance, out int payloadLength)
                || recordType != OfficeArtFbse
                || !OfficeArtBlipStoreEntryReader.TryRead(
                    data,
                    blipStoreOffset + 8,
                    payloadLength,
                    recordInstance,
                    delayStream: null,
                    out OfficeArtBlipStoreEntry? entry,
                    maximumDecodedImageBytes)
                || entry?.HasImportableImage != true) {
                warning = $"The picture data at DOC Data stream offset {offset} has no importable embedded OfficeArt BLIP.";
                return false;
            }

            double widthPixels = widthGoal * (widthScale / 1000D) * PixelsPerTwip;
            double heightPixels = heightGoal * (heightScale / 1000D) * PixelsPerTwip;
            picture = new LegacyDocPicture(entry.ImageBytes, entry.ContentType!, widthPixels, heightPixels);
            consumedLength = lcb;
            return true;
        }

        private static bool TryReadOfficeArtHeader(
            byte[] data,
            int offset,
            int boundary,
            out ushort recordType,
            out ushort recordInstance,
            out int payloadLength) {
            recordType = 0;
            recordInstance = 0;
            payloadLength = 0;
            if (offset < 0 || offset > boundary - 8 || boundary > data.Length) {
                return false;
            }

            ushort versionAndInstance = LegacyDocFib.ReadUInt16(data, offset);
            recordInstance = unchecked((ushort)(versionAndInstance >> 4));
            recordType = LegacyDocFib.ReadUInt16(data, offset + 2);
            uint declaredLength = unchecked((uint)LegacyDocFib.ReadInt32(data, offset + 4));
            if (declaredLength > int.MaxValue || declaredLength > unchecked((uint)(boundary - offset - 8))) {
                return false;
            }

            payloadLength = checked((int)declaredLength);
            return true;
        }

        private static LegacyDocCharacterFormat GetFormatForFileOffset(
            IReadOnlyList<LegacyDocCharacterFormatRange> ranges,
            int fileOffset) {
            for (int index = 0; index < ranges.Count; index++) {
                if (ranges[index].Contains(fileOffset)) {
                    return ranges[index].Format;
                }
            }

            return LegacyDocCharacterFormat.Default;
        }

        private static bool CoversDataStream(IReadOnlyDictionary<int, int> ranges, int dataLength) {
            int position = 0;
            foreach (KeyValuePair<int, int> range in ranges.OrderBy(item => item.Key)) {
                if (range.Key != position || range.Value <= 0) {
                    return false;
                }

                position = checked(position + range.Value);
            }

            return position == dataLength;
        }

        internal readonly struct LegacyDocPictureReadResult {
            internal static LegacyDocPictureReadResult Empty { get; } = new(
                new Dictionary<int, LegacyDocPicture>(),
                fullyProjectsDataStream: false,
                warning: null);

            internal LegacyDocPictureReadResult(
                IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition,
                bool fullyProjectsDataStream,
                string? warning) {
                PicturesByCharacterPosition = picturesByCharacterPosition;
                FullyProjectsDataStream = fullyProjectsDataStream;
                Warning = warning;
            }

            internal IReadOnlyDictionary<int, LegacyDocPicture> PicturesByCharacterPosition { get; }
            internal bool FullyProjectsDataStream { get; }
            internal string? Warning { get; }
        }
    }
}
