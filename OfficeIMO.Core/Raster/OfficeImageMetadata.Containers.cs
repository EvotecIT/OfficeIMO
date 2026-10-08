using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Drawing;

/// <summary>Encoded profile families supported by lossless image metadata removal.</summary>
[Flags]
public enum OfficeImageMetadataProfileKinds {
    /// <summary>No profile families.</summary>
    None = 0,
    /// <summary>Exif profiles.</summary>
    Exif = 1,
    /// <summary>Standard and extended XMP packets.</summary>
    Xmp = 2,
    /// <summary>ICC color profiles.</summary>
    Icc = 4,
    /// <summary>IPTC IIM image resource data.</summary>
    Iptc = 8,
    /// <summary>Embedded C2PA manifest carriers.</summary>
    C2pa = 16,
    /// <summary>All supported profile families.</summary>
    All = Exif | Xmp | Icc | Iptc | C2pa
}

/// <summary>Lossless encoded-profile removal evidence.</summary>
public sealed class OfficeImageMetadataRemovalResult {
    private readonly byte[] _encoded;
    internal OfficeImageMetadataRemovalResult(byte[] encoded, OfficeImageMetadataProfileKinds present, OfficeImageMetadataProfileKinds removed) { _encoded = encoded; PresentProfiles = present; RemovedProfiles = removed; }
    /// <summary>A copy of the rewritten encoded image.</summary>
    public byte[] EncodedBytes => (byte[])_encoded.Clone();
    /// <summary>Profile families found in the input.</summary>
    public OfficeImageMetadataProfileKinds PresentProfiles { get; }
    /// <summary>Profile families removed from the input.</summary>
    public OfficeImageMetadataProfileKinds RemovedProfiles { get; }
}

public sealed partial class OfficeImageMetadata {
    private static readonly byte[] JpegExifPrefix = Encoding.ASCII.GetBytes("Exif\0\0");
    private static readonly byte[] JpegXmpPrefix = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
    private static readonly byte[] JpegExtendedXmpPrefix = Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
    private static readonly byte[] JpegIccPrefix = Encoding.ASCII.GetBytes("ICC_PROFILE\0");
    private static readonly byte[] PhotoshopPrefix = Encoding.ASCII.GetBytes("Photoshop 3.0\0");

    /// <summary>Removes selected supported profile families while preserving encoded image payloads, frame timing, and TIFF pages.</summary>
    public static OfficeImageMetadataRemovalResult Remove(byte[] encodedBytes, OfficeImageMetadataProfileKinds kinds, CancellationToken cancellationToken = default) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        if ((kinds & ~OfficeImageMetadataProfileKinds.All) != 0) throw new ArgumentOutOfRangeException(nameof(kinds));
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(encodedBytes.Length)) throw new FormatException("Image bytes exceed the metadata-edit limit.");
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info)) throw new FormatException("The image container is malformed or unsupported.");
        byte[] working = encodedBytes;
        OfficeImageMetadataProfileKinds present = OfficeImageMetadataProfileKinds.None;
        OfficeImageMetadataProfileKinds removed = OfficeImageMetadataProfileKinds.None;
        OfficeProvenanceReport provenance = OfficeProvenanceInspector.Inspect(encodedBytes, options: new OfficeProvenanceOptions { MaxAssetBytes = OfficeRasterGuards.MaximumEncodedBytes, CancellationToken = cancellationToken });
        foreach (OfficeProvenanceEvidence evidence in provenance.Evidence) if (evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest) present |= OfficeImageMetadataProfileKinds.C2pa;
        OfficeImageMetadataProfileKinds originalC2pa = present;
        if ((kinds & originalC2pa & OfficeImageMetadataProfileKinds.C2pa) != 0) {
            if (checked(encodedBytes.LongLength * 5L + provenance.ExpandedInspectionBytes) > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentException("Metadata rewriting exceeds the managed working-set limit.");
            var options = new OfficeProvenanceRemovalOptions { RemoveC2paManifests = true, RemoveExternalC2paReferences = false, RemoveAiSourceMetadata = false, RequireStructurallyValidCarrier = true };
            options.Limits.MaxAssetBytes = OfficeRasterGuards.MaximumEncodedBytes;
            options.Limits.CancellationToken = cancellationToken;
            OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(working, options: options);
            foreach (OfficeProvenanceChange change in result.Changes) if (change.Carrier == OfficeProvenanceCarrierKind.C2paManifest) present |= removed |= OfficeImageMetadataProfileKinds.C2pa;
            working = result.ToArray();
        }
        OfficeImageMetadataProfileKinds profileKinds = kinds & ~OfficeImageMetadataProfileKinds.C2pa;
        long retainedInput = ReferenceEquals(working, encodedBytes) ? 0L : encodedBytes.LongLength;
        byte[] output = info.Format switch {
            OfficeImageFormat.Jpeg => RewriteJpeg(working, null, cancellationToken, profileKinds, out present, retainedInput),
            OfficeImageFormat.Png => RewritePng(working, null, cancellationToken, profileKinds, out present, retainedInput),
            OfficeImageFormat.Webp => RewriteWebp(working, null, cancellationToken, profileKinds, out present, retainedInput),
            OfficeImageFormat.Tiff => RemoveTiffProfiles(working, profileKinds, cancellationToken, out present, retainedInput),
            OfficeImageFormat.Bmp => RewriteBmp(working, null, profileKinds, cancellationToken, out present, retainedInput),
            OfficeImageFormat.Gif => RewriteGif(working, null, profileKinds, cancellationToken, out present, retainedInput),
            OfficeImageFormat.PortableMap or OfficeImageFormat.Tga => CloneWithoutProfiles(working, info.Format, cancellationToken, out present),
            _ => throw new NotSupportedException("Lossless profile removal is not supported for this image container.")
        };
        removed |= present & profileKinds;
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeImageMetadataRemovalResult(output, present | removed | originalC2pa, removed);
    }

    private static byte[] RewriteJpeg(byte[] input, OfficeImageMetadata metadata, CancellationToken token) => RewriteJpeg(input, metadata, token, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _);
    private static byte[] RewriteJpeg(byte[] input, OfficeImageMetadata? metadata, CancellationToken token, OfficeImageMetadataProfileKinds replace, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        present = OfficeImageMetadataProfileKinds.None;
        using var output = CreateRewriteStream(input, metadata, token, additionallyRetainedBytes); output.Write(input, 0, 2);
        // Insert the new resource once before copying existing segments. This avoids
        // materializing and then copying an entire completed JPEG to add IPTC.
        bool iptcWritten = metadata?._iptc != null;
        if (iptcWritten) WriteJpegSegment(output, 0xED, RewritePhotoshop(PhotoshopPrefix, true, metadata!._iptc, out _));
        if (metadata != null) {
            byte[]? exif = metadata.HasExifProfile ? EncodeDensityExif(metadata, token) : null; if (exif != null) WriteJpegSegment(output, 0xE1, Join(JpegExifPrefix, exif));
            if (!HasJpegJfif(input)) {
                GetExifResolution(metadata, out double x, out double y, out ushort unit);
                byte[] jfif = new byte[] { 74, 70, 73, 70, 0, 1, 2, unit == 1 ? (byte)0 : unit == 3 ? (byte)2 : (byte)1, 0, 0, 0, 0, 0, 0 };
                OfficeExifProfileCodec.Write(jfif, 8, Density(x, ushort.MaxValue), 2, false);
                OfficeExifProfileCodec.Write(jfif, 10, Density(y, ushort.MaxValue), 2, false);
                WriteJpegSegment(output, 0xE0, jfif);
            }
            if (metadata._xmp != null) WriteJpegSegment(output, 0xE1, Join(JpegXmpPrefix, metadata._xmp));
            if (metadata._icc != null) {
                const int blockSize = 65519; int blocks = (metadata._icc.Length + blockSize - 1) / blockSize;
                if (blocks > 255) throw new FormatException("ICC profile requires too many JPEG segments.");
                for (int index = 0; index < blocks; index++) {
                    int size = Math.Min(blockSize, metadata._icc.Length - index * blockSize); byte[] payload = new byte[JpegIccPrefix.Length + 2 + size];
                    Buffer.BlockCopy(JpegIccPrefix, 0, payload, 0, JpegIccPrefix.Length); payload[JpegIccPrefix.Length] = (byte)(index + 1); payload[JpegIccPrefix.Length + 1] = (byte)blocks;
                    Buffer.BlockCopy(metadata._icc, index * blockSize, payload, JpegIccPrefix.Length + 2, size); WriteJpegSegment(output, 0xE2, payload);
                }
            }
        }
        int cursor = 2; bool inScan = false;
        while (cursor < input.Length) {
            token.ThrowIfCancellationRequested();
            if (inScan) {
                int start = cursor;
                while (cursor < input.Length) {
                    if ((cursor & 4095) == 0) token.ThrowIfCancellationRequested();
                    if (input[cursor] != 0xFF) { cursor++; continue; }
                    int next = cursor + 1; while (next < input.Length && input[next] == 0xFF) next++;
                    if (next >= input.Length) throw new FormatException("Truncated JPEG entropy marker.");
                    if (input[next] == 0 || input[next] >= 0xD0 && input[next] <= 0xD7) { cursor = next + 1; continue; }
                    break;
                }
                output.Write(input, start, cursor - start); inScan = false;
                if (cursor == input.Length) throw new FormatException("JPEG is missing its end marker.");
            }
            int segmentStart = cursor;
            if (!OfficeProvenanceJpeg.TryReadMarker(input, cursor, out byte marker, out int payload, out int length, out int end)) throw new FormatException("JPEG contains a malformed segment.");
            OfficeImageMetadataProfileKinds kind = marker == 0xE1 && PrefixAt(input, payload, length, JpegExifPrefix) ? OfficeImageMetadataProfileKinds.Exif :
                marker == 0xE1 && (PrefixAt(input, payload, length, JpegXmpPrefix) || PrefixAt(input, payload, length, JpegExtendedXmpPrefix)) ? OfficeImageMetadataProfileKinds.Xmp :
                marker == 0xE2 && PrefixAt(input, payload, length, JpegIccPrefix) ? OfficeImageMetadataProfileKinds.Icc : OfficeImageMetadataProfileKinds.None;
            present |= kind;
            if (marker == 0xED && PrefixAt(input, payload, length, PhotoshopPrefix)) {
                byte[] rewritten = RewritePhotoshop(Slice(input, payload, length), (replace & OfficeImageMetadataProfileKinds.Iptc) != 0, iptcWritten ? null : metadata?._iptc, out bool found);
                if (found) present |= OfficeImageMetadataProfileKinds.Iptc;
                if ((replace & OfficeImageMetadataProfileKinds.Iptc) != 0) { if (rewritten.Length > PhotoshopPrefix.Length) WriteJpegSegment(output, marker, rewritten); iptcWritten |= metadata?._iptc != null; }
                else output.Write(input, segmentStart, end - segmentStart);
            } else if ((replace & kind) == 0 || kind == OfficeImageMetadataProfileKinds.None) {
                if (metadata != null && marker == 0xE0 && PrefixAt(input, payload, length, Encoding.ASCII.GetBytes("JFIF\0")) && length >= 12) {
                    byte[] copy = Slice(input, segmentStart, end - segmentStart); int at = payload - segmentStart;
                    GetExifResolution(metadata, out double x, out double y, out ushort unit);
                    copy[at + 7] = unit == 1 ? (byte)0 : unit == 3 ? (byte)2 : (byte)1;
                    OfficeExifProfileCodec.Write(copy, at + 8, Density(x, ushort.MaxValue), 2, false);
                    OfficeExifProfileCodec.Write(copy, at + 10, Density(y, ushort.MaxValue), 2, false); output.Write(copy, 0, copy.Length);
                } else output.Write(input, segmentStart, end - segmentStart);
            }
            cursor = end;
            if (marker == 0xDA) inScan = true;
            if (marker == 0xD9) { if (cursor != input.Length) output.Write(input, cursor, input.Length - cursor); break; }
            if (output.Length > OfficeRasterGuards.MaximumEncodedBytes) throw new FormatException("Edited JPEG exceeds the encoded-size limit.");
        }
        return output.ToArray();
    }

    private static byte[]? ReadJpegIptc(byte[] input, CancellationToken token) {
        int cursor = 2;
        while (cursor < input.Length && OfficeProvenanceJpeg.TryReadMarker(input, cursor, out byte marker, out int payload, out int length, out int end)) {
            token.ThrowIfCancellationRequested();
            if (marker == 0xDA || marker == 0xD9) break;
            if (marker == 0xED && PrefixAt(input, payload, length, PhotoshopPrefix)) {
                foreach (Resource resource in PhotoshopResources(Slice(input, payload, length))) if (resource.Id == 0x0404) return resource.Payload;
            }
            cursor = end;
        }
        return null;
    }
    private sealed class Resource { internal ushort Id; internal byte[] Encoded = Array.Empty<byte>(); internal byte[] Payload = Array.Empty<byte>(); }
    private static IEnumerable<Resource> PhotoshopResources(byte[] input) {
        int cursor = PhotoshopPrefix.Length;
        while (cursor < input.Length) {
            int start = cursor;
            if (cursor > input.Length - 7 || !PrefixAt(input, cursor, input.Length - cursor, Encoding.ASCII.GetBytes("8BIM"))) throw new FormatException("Malformed Photoshop image resource.");
            ushort id = (ushort)OfficeExifProfileCodec.Read(input, cursor + 4, 2, false); cursor += 6;
            int nameLength = input[cursor]; cursor += 1 + nameLength; if (((1 + nameLength) & 1) != 0) cursor++;
            if (cursor > input.Length - 4) throw new FormatException("Truncated Photoshop image resource.");
            uint sizeValue = (uint)OfficeExifProfileCodec.Read(input, cursor, 4, false); cursor += 4;
            if (sizeValue > int.MaxValue || sizeValue > input.Length - cursor) throw new FormatException("Photoshop image resource is outside the segment.");
            int size = (int)sizeValue; byte[] payload = Slice(input, cursor, size); cursor += size + (size & 1);
            if (cursor > input.Length) throw new FormatException("Missing Photoshop image resource padding.");
            yield return new Resource { Id = id, Payload = payload, Encoded = Slice(input, start, cursor - start) };
        }
    }
    private static byte[] RewritePhotoshop(byte[] input, bool removeIptc, byte[]? replacement, out bool found) {
        found = false; using var output = new MemoryStream(); output.Write(PhotoshopPrefix, 0, PhotoshopPrefix.Length);
        foreach (Resource resource in PhotoshopResources(input)) { if (resource.Id == 0x0404) { found = true; if (removeIptc) continue; } output.Write(resource.Encoded, 0, resource.Encoded.Length); }
        if (removeIptc && replacement != null) {
            var header = new byte[] { 56, 66, 73, 77, 4, 4, 0, 0, 0, 0, 0, 0 }; OfficeExifProfileCodec.Write(header, 8, (uint)replacement.Length, 4, false);
            output.Write(header, 0, header.Length); output.Write(replacement, 0, replacement.Length); if ((replacement.Length & 1) != 0) output.WriteByte(0);
        }
        return output.ToArray();
    }

    private static byte[] RewritePng(byte[] input, OfficeImageMetadata metadata, CancellationToken token) => RewritePng(input, metadata, token, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _);
    private static byte[] RewritePng(byte[] input, OfficeImageMetadata? metadata, CancellationToken token, OfficeImageMetadataProfileKinds replace, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        if (metadata?._iptc != null) throw new NotSupportedException("PNG does not have a standard IPTC IIM profile chunk.");
        present = OfficeImageMetadataProfileKinds.None;
        using var output = CreateRewriteStream(input, metadata, token, additionallyRetainedBytes); output.Write(input, 0, 8); int cursor = 8;
        while (cursor <= input.Length - 12) {
            token.ThrowIfCancellationRequested(); int length = checked((int)OfficeExifProfileCodec.Read(input, cursor, 4, false));
            if (length > input.Length - cursor - 12) throw new FormatException("PNG metadata chunk is truncated.");
            string type = Encoding.ASCII.GetString(input, cursor + 4, 4);
            OfficeImageMetadataProfileKinds kind = type == "eXIf" ? OfficeImageMetadataProfileKinds.Exif : type == "iCCP" ? OfficeImageMetadataProfileKinds.Icc : IsPngXmp(input, cursor + 8, length, type) ? OfficeImageMetadataProfileKinds.Xmp : type == "caBX" ? OfficeImageMetadataProfileKinds.C2pa : OfficeImageMetadataProfileKinds.None;
            present |= kind;
            bool skip = (replace & kind) != 0 || metadata != null && (type == "pHYs" || metadata._icc != null && type == "sRGB");
            if (!skip) output.Write(input, cursor, length + 12);
            if (metadata != null && type == "IHDR") {
                byte[]? exif = metadata.EncodeExifProfile(token); if (exif != null) WritePngChunk(output, "eXIf", exif);
                if (metadata._icc != null) { byte[] compressed = OfficeZlibCodec.Compress(metadata._icc, token); WritePngChunk(output, "iCCP", Join(new byte[] { 73, 67, 67, 0, 0 }, compressed)); }
                if (metadata._xmp != null) WritePngChunk(output, "iTXt", Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp\0\0\0\0\0"), metadata._xmp));
                byte[] density = new byte[9]; double scale = metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerInch ? 1D / 0.0254D : metadata.ResolutionUnits == OfficeImageResolutionUnit.PixelsPerCentimeter ? 100D : 1D;
                OfficeExifProfileCodec.Write(density, 0, Density(metadata.HorizontalResolution * scale, uint.MaxValue), 4, false); OfficeExifProfileCodec.Write(density, 4, Density(metadata.VerticalResolution * scale, uint.MaxValue), 4, false);
                density[8] = metadata.ResolutionUnits == OfficeImageResolutionUnit.AspectRatio ? (byte)0 : (byte)1; WritePngChunk(output, "pHYs", density);
            }
            cursor += length + 12;
            if (output.Length > OfficeRasterGuards.MaximumEncodedBytes) throw new FormatException("Edited PNG exceeds the encoded-size limit.");
        }
        if (cursor != input.Length) throw new FormatException("PNG contains truncated trailing data.");
        return output.ToArray();
    }

    private static void ReadPngProfiles(byte[] input, OfficeImageMetadata metadata, CancellationToken token) {
        for (int cursor = 8; cursor <= input.Length - 12;) {
            token.ThrowIfCancellationRequested(); int length = checked((int)OfficeExifProfileCodec.Read(input, cursor, 4, false)); string type = Encoding.ASCII.GetString(input, cursor + 4, 4); int payload = cursor + 8;
            if (type == "eXIf") metadata.SetExifProfile(Slice(input, payload, length), token);
            if (type == "iCCP") { int end = Array.IndexOf(input, (byte)0, payload, length); if (end < 0 || end + 2 > payload + length) throw new FormatException("Invalid PNG ICC profile framing."); metadata.IccProfile = OfficeZlibCodec.Decompress(Slice(input, end + 2, payload + length - end - 2), OfficeExifProfileCodec.MaximumProfileBytes, cancellationToken: token); }
            if (IsPngXmp(input, payload, length, type)) {
                int end = Array.IndexOf(input, (byte)0, payload, length); int text = end + 1; bool compressed = type == "zTXt";
                if (type == "zTXt") text++;
                if (type == "iTXt") { compressed = input[text] == 1; text += 2; int languageEnd = Array.IndexOf(input, (byte)0, text, payload + length - text); if (languageEnd < 0) throw new FormatException("Invalid PNG XMP language field."); text = languageEnd + 1; int translatedEnd = Array.IndexOf(input, (byte)0, text, payload + length - text); if (translatedEnd < 0) throw new FormatException("Invalid PNG XMP keyword field."); text = translatedEnd + 1; }
                byte[] xmp = Slice(input, text, payload + length - text); metadata.XmpProfile = compressed ? OfficeZlibCodec.Decompress(xmp, OfficeExifProfileCodec.MaximumProfileBytes, cancellationToken: token) : xmp;
            }
            cursor += length + 12;
        }
    }
    private static bool IsPngXmp(byte[] input, int offset, int length, string type) => (type == "iTXt" || type == "tEXt" || type == "zTXt") && PrefixAt(input, offset, length, Encoding.ASCII.GetBytes("XML:com.adobe.xmp\0"));
    private static void WritePngChunk(Stream output, string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12]; OfficeExifProfileCodec.Write(chunk, 0, (uint)payload.Length, 4, false); Buffer.BlockCopy(Encoding.ASCII.GetBytes(type), 0, chunk, 4, 4); Buffer.BlockCopy(payload, 0, chunk, 8, payload.Length);
        OfficeExifProfileCodec.Write(chunk, chunk.Length - 4, OfficePngCrc32.Compute(chunk, 4, payload.Length + 4), 4, false); output.Write(chunk, 0, chunk.Length);
    }

    private static byte[] RewriteWebp(byte[] input, OfficeImageMetadata metadata, CancellationToken token) => RewriteWebp(input, metadata, token, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _);
    private static byte[] RewriteWebp(byte[] input, OfficeImageMetadata? metadata, CancellationToken token, OfficeImageMetadataProfileKinds replace, out OfficeImageMetadataProfileKinds present, long additionallyRetainedBytes = 0L) {
        if (metadata?._iptc != null) throw new NotSupportedException("WebP does not have an IPTC IIM profile chunk.");
        present = OfficeImageMetadataProfileKinds.None;
        using var output = CreateRewriteStream(input, metadata, token, additionallyRetainedBytes);
        byte[]? exif = null;
        if (metadata != null) exif = EncodeDensityExif(metadata, token);
        bool hasExtended = input.Length >= 30 && Encoding.ASCII.GetString(input, 12, 4) == "VP8X";
        output.Write(Encoding.ASCII.GetBytes("RIFF\0\0\0\0WEBP"), 0, 12);
        bool profiles = exif != null || metadata?._xmp != null || metadata?._icc != null;
        if (!hasExtended && profiles) {
            OfficeImageInfo info = OfficeImageReader.Identify(input); byte[] header = new byte[10];
            header[0] = (byte)((exif != null ? 8 : 0) | (metadata?._xmp != null ? 4 : 0) | (metadata?._icc != null ? 32 : 0));
            if (input.Length >= 25 && Encoding.ASCII.GetString(input, 12, 4) == "VP8L" && (input[24] & 0x10) != 0) header[0] |= 16;
            OfficeExifProfileCodec.Write(header, 4, (uint)(info.Width - 1), 3, true); OfficeExifProfileCodec.Write(header, 7, (uint)(info.Height - 1), 3, true); WriteWebpChunk(output, "VP8X", header);
        }
        if (!hasExtended && metadata?._icc != null) WriteWebpChunk(output, "ICCP", metadata._icc);
        for (int cursor = 12; cursor <= input.Length - 8;) {
            token.ThrowIfCancellationRequested(); int length = checked((int)OfficeExifProfileCodec.Read(input, cursor + 4, 4, true)); if (length > input.Length - cursor - 8) throw new FormatException("Truncated WebP metadata chunk."); string type = Encoding.ASCII.GetString(input, cursor, 4);
            OfficeImageMetadataProfileKinds kind = type == "EXIF" ? OfficeImageMetadataProfileKinds.Exif : type == "XMP " ? OfficeImageMetadataProfileKinds.Xmp : type == "ICCP" ? OfficeImageMetadataProfileKinds.Icc : OfficeImageMetadataProfileKinds.None; present |= kind;
            if (type == "VP8X") {
                hasExtended = true; byte[] payload = Slice(input, cursor + 8, length); if (length != 10) throw new FormatException("Invalid extended WebP header.");
                byte clear = 0; if ((replace & OfficeImageMetadataProfileKinds.Exif) != 0) clear |= 8; if ((replace & OfficeImageMetadataProfileKinds.Xmp) != 0) clear |= 4; if ((replace & OfficeImageMetadataProfileKinds.Icc) != 0) clear |= 32;
                payload[0] &= (byte)~clear; if (exif != null) payload[0] |= 8; if (metadata?._xmp != null) payload[0] |= 4; if (metadata?._icc != null) payload[0] |= 32; WriteWebpChunk(output, type, payload);
                if (metadata?._icc != null) WriteWebpChunk(output, "ICCP", metadata._icc);
            } else if ((replace & kind) == 0 || kind == OfficeImageMetadataProfileKinds.None) output.Write(input, cursor, length + 8 + (length & 1));
            cursor += length + 8 + (length & 1);
        }
        if (exif != null) WriteWebpChunk(output, "EXIF", exif); if (metadata?._xmp != null) WriteWebpChunk(output, "XMP ", metadata._xmp);
        byte[] result = output.ToArray(); OfficeExifProfileCodec.Write(result, 4, (uint)(result.Length - 8), 4, true); return result;
    }
    private static void ReadWebpProfiles(byte[] input, OfficeImageMetadata metadata, CancellationToken token) {
        for (int cursor = 12; cursor <= input.Length - 8;) {
            int length = checked((int)OfficeExifProfileCodec.Read(input, cursor + 4, 4, true)); string type = Encoding.ASCII.GetString(input, cursor, 4);
            token.ThrowIfCancellationRequested();
            if (type == "EXIF") metadata.SetExifProfile(Slice(input, cursor + 8, length), token); if (type == "XMP ") metadata.XmpProfile = Slice(input, cursor + 8, length); if (type == "ICCP") metadata.IccProfile = Slice(input, cursor + 8, length); cursor += length + 8 + (length & 1);
        }
    }
    private static void WriteWebpChunk(Stream output, string type, byte[] payload) { byte[] header = new byte[8]; Buffer.BlockCopy(Encoding.ASCII.GetBytes(type), 0, header, 0, 4); OfficeExifProfileCodec.Write(header, 4, (uint)payload.Length, 4, true); output.Write(header, 0, 8); output.Write(payload, 0, payload.Length); if ((payload.Length & 1) != 0) output.WriteByte(0); }
    private static void WriteJpegSegment(Stream output, byte marker, byte[] payload) { if (payload.Length > 65533) throw new FormatException("Metadata exceeds the JPEG segment size limit."); output.WriteByte(255); output.WriteByte(marker); int length = payload.Length + 2; output.WriteByte((byte)(length >> 8)); output.WriteByte((byte)length); output.Write(payload, 0, payload.Length); }
    private static uint Density(double value, uint maximum) { if (value < 1D || value > maximum) throw new ArgumentOutOfRangeException(nameof(value), "Resolution is outside the container's representable range."); return checked((uint)Math.Round(value, MidpointRounding.AwayFromZero)); }
    private static byte[] Join(byte[] prefix, byte[] payload) { var result = new byte[checked(prefix.Length + payload.Length)]; Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length); Buffer.BlockCopy(payload, 0, result, prefix.Length, payload.Length); return result; }
    private static byte[] Slice(byte[] bytes, int offset, int count) { var result = new byte[count]; Buffer.BlockCopy(bytes, offset, result, 0, count); return result; }
    private static bool StartsWith(byte[] input, byte[] prefix) => PrefixAt(input, 0, input.Length, prefix);
    private static bool PrefixAt(byte[] input, int offset, int count, byte[] prefix) { if (count < prefix.Length) return false; for (int i = 0; i < prefix.Length; i++) if (input[offset + i] != prefix[i]) return false; return true; }
}
