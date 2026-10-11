using System;
using System.Threading;

namespace OfficeIMO.Drawing {
    public static partial class OfficeHeifMetadataReader {
        /// <summary>Replaces or clears selected existing HEIF EXIF and XMP profiles without decoding pixels.</summary>
        /// <param name="data">Encoded input bytes, which remain unchanged.</param>
        /// <param name="metadata">Replacement profiles. Null clears every selected profile; a missing selected profile also clears its item.</param>
        /// <param name="profiles">The EXIF and XMP families to edit. None returns an independent unchanged copy.</param>
        /// <param name="output">Independently owned complete output on success, or null when any selected edit is rejected.</param>
        /// <param name="cancellationToken">Cancellation observed during parsing, encoding, and copying.</param>
        /// <returns>True only when every selected edit completes within the resource limits.</returns>
        /// <remarks>
        /// Unselected profiles remain untouched. Every selected family must have an existing,
        /// uniquely selected writable item; missing, protected, encoded, externally located,
        /// overlapping, item-data-box, and multiple-extent items are not created or rewritten.
        /// XMP replacement bytes must be valid UTF-8. Malformed existing payloads may be replaced
        /// without decoding them. No intermediate output is published; both edits share parser
        /// work limits and account for the original input, supplied profiles, and retained output.
        /// The encoded ceiling is 128 MiB, metadata payload ceiling 16 MiB, and managed working-set
        /// ceiling 256 MiB. Cancellation throws rather than returning a partial result.
        /// </remarks>
        /// <exception cref="ArgumentNullException">Data is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Profiles contains a family other than EXIF or XMP.</exception>
        /// <exception cref="OperationCanceledException">The operation is cancelled.</exception>
        public static bool TryWriteMetadata(byte[] data, OfficeImageMetadata? metadata,
            OfficeImageMetadataProfileKinds profiles, out byte[]? output, CancellationToken cancellationToken = default) {
            output = null;
            if ((profiles & ~(OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp)) != 0) {
                throw new ArgumentOutOfRangeException(nameof(profiles));
            }
            byte[]? value = null;
            bool success = TryRun(data, cancellationToken,
                parser => parser.TryWriteMetadata(data, metadata, profiles, out value), metadata?.RetainedProfileBytes ?? 0L);
            output = success ? value : null;
            return success;
        }

        private sealed partial class Parser {
            internal bool TryWriteMetadata(byte[] data, OfficeImageMetadata? metadata,
                OfficeImageMetadataProfileKinds profiles, out byte[]? result) {
                result = null;
                if (profiles == OfficeImageMetadataProfileKinds.None) {
                    ReserveBytes(data.LongLength + 24L);
                    byte[] copy = new byte[data.Length];
                    CopyBytes(data, 0, copy, 0, data.Length);
                    result = copy;
                    return true;
                }
                byte[] working = data;
                bool writeExif = (profiles & OfficeImageMetadataProfileKinds.Exif) != 0;
                if (writeExif) {
                    if (!TryWriteExifProfile(working, metadata, out byte[]? exifOutput)) {
                        return false;
                    }
                    working = exifOutput!;
                }
                if ((profiles & OfficeImageMetadataProfileKinds.Xmp) != 0) {
                    if (writeExif) {
                        // Original input and the first output remain live through the second edit.
                        ReserveBytes(working.LongLength + 24L);
                    }
                    if (!TryWriteXmpProfile(working, metadata?.XmpProfileData, out byte[]? xmpOutput)) {
                        return false;
                    }
                    working = xmpOutput!;
                }
                _cancellationToken.ThrowIfCancellationRequested();
                result = working;
                return true;
            }

            private bool TryWriteXmpProfile(byte[] data, byte[]? packet, out byte[]? result) {
                result = null;
                if (!TryFindWritableMetadataItem(data, exif: false, out IlocItem item, out ItemExtent extent)) {
                    return false;
                }
                if (packet != null) {
                    if (packet.Length > OfficeExifProfileCodec.MaximumProfileBytes) {
                        return false;
                    }
                    _cancellationToken.ThrowIfCancellationRequested();
                    _ = StrictXmpUtf8.GetCharCount(packet);
                    _cancellationToken.ThrowIfCancellationRequested();
                }
                return TryWriteItemData(data, item, extent, out result, packet ?? Array.Empty<byte>());
            }
        }
    }
}
