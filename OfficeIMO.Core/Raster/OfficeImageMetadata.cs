using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>The unit used by image resolution values.</summary>
public enum OfficeImageResolutionUnit {
    /// <summary>Unitless horizontal and vertical aspect ratio.</summary>
    AspectRatio,
    /// <summary>Pixels per inch.</summary>
    PixelsPerInch,
    /// <summary>Pixels per centimeter.</summary>
    PixelsPerCentimeter,
    /// <summary>Pixels per meter, used by PNG and Windows bitmap density fields.</summary>
    PixelsPerMeter
}

/// <summary>Portable image metadata with typed Exif edits and independently owned profile bytes.</summary>
/// <remarks>Metadata edits do not apply ICC color transformations to image pixels. Use <see cref="Apply"/> to replace supported container metadata without recompressing image data.</remarks>
public sealed partial class OfficeImageMetadata {
    private OfficeExifProfileCodec.Profile? _exif;
    private readonly Dictionary<OfficeExifTag, OfficeExifValue> _changes = new Dictionary<OfficeExifTag, OfficeExifValue>();
    private readonly HashSet<OfficeExifTag> _removed = new HashSet<OfficeExifTag>();
    private byte[]? _xmp;
    private byte[]? _icc;
    private byte[]? _iptc;
    private readonly HashSet<OfficeExifTag> _tiffOpaqueOffsets = new HashSet<OfficeExifTag>();
    private OfficeImageResolution _resolution = new OfficeImageResolution(96D, 96D);

    /// <summary>Horizontal resolution in <see cref="ResolutionUnits"/>.</summary>
    public double HorizontalResolution { get => _resolution.Horizontal; set => Resolution = new OfficeImageResolution(value, _resolution.Vertical, _resolution.Unit); }
    /// <summary>Vertical resolution in <see cref="ResolutionUnits"/>.</summary>
    public double VerticalResolution { get => _resolution.Vertical; set => Resolution = new OfficeImageResolution(_resolution.Horizontal, value, _resolution.Unit); }
    /// <summary>The unit for both resolution values.</summary>
    public OfficeImageResolutionUnit ResolutionUnits { get => _resolution.Unit; set => Resolution = new OfficeImageResolution(_resolution.Horizontal, _resolution.Vertical, value); }
    /// <summary>The authoritative native density or unitless aspect ratio, suitable for encoder overrides.</summary>
    /// <remarks>Assigning resolution explicitly restores its Exif density fields after previous individual removals.
    /// When Exif density fields are present or removed, the new values must fit Exif's unsigned rational representation.
    /// Rejected assignments retain the previous resolution, fields, and removal intent.</remarks>
    public OfficeImageResolution Resolution {
        get => _resolution;
        set {
            SetResolutionAndDensityFields(value ?? throw new ArgumentNullException(nameof(value)));
            _removed.Remove(OfficeExifTag.XResolution);
            _removed.Remove(OfficeExifTag.YResolution);
            _removed.Remove(OfficeExifTag.ResolutionUnit);
        }
    }
    /// <summary>Horizontal physical resolution in dots per inch, or null for a unitless aspect ratio.</summary>
    public double? PhysicalDpiX => ToPhysicalDpi(HorizontalResolution);
    /// <summary>Vertical physical resolution in dots per inch, or null for a unitless aspect ratio.</summary>
    public double? PhysicalDpiY => ToPhysicalDpi(VerticalResolution);
    /// <summary>Whether any Exif fields or an opaque Exif profile are present.</summary>
    public bool HasExifProfile => _exif != null || _changes.Count != 0;
    /// <summary>Whether TIFF-relative opaque fields require their original TIFF container for safe preservation.</summary>
    public bool RequiresOriginalTiffContainer => _tiffOpaqueOffsets.Count != 0;
    /// <summary>Classic TIFF Exif profile without a JPEG Exif prefix. Setting a profile replaces all pending Exif edits.</summary>
    public byte[]? ExifProfile {
        get => EncodeExifProfile();
        set => SetExifProfile(value, CancellationToken.None);
    }
    /// <summary>XMP XML packet, excluding container framing. Profile arrays are copied.</summary>
    public byte[]? XmpProfile { get => Copy(_xmp); set => _xmp = BoundedCopy(value); }
    /// <summary>ICC profile bytes. Profile arrays are copied.</summary>
    public byte[]? IccProfile { get => Copy(_icc); set => _icc = BoundedCopy(value); }
    /// <summary>IPTC IIM data, excluding Photoshop resource framing. Profile arrays are copied.</summary>
    public byte[]? IptcProfile { get => Copy(_iptc); set => _iptc = BoundedCopy(value); }

    /// <summary>Current Exif values from the image, camera, GPS, and interoperability directories.</summary>
    public IReadOnlyList<OfficeExifValue> ExifValues {
        get {
            var result = new List<OfficeExifValue>();
            var seen = new HashSet<OfficeExifTag>();
            if (_exif != null) foreach (OfficeExifProfileCodec.Directory directory in _exif.Directories.Values) foreach (OfficeExifProfileCodec.Field entry in directory.Fields) {
                if (IsStructural(entry.Tag.Id) || _removed.Contains(entry.Tag) || !seen.Add(entry.Tag)) continue;
                result.Add(_changes.TryGetValue(entry.Tag, out OfficeExifValue? changed) ? changed : new OfficeExifValue(entry, _exif.Bytes, _exif.Little, copyEncoding: false));
            }
            foreach (KeyValuePair<OfficeExifTag, OfficeExifValue> change in _changes) if (seen.Add(change.Key)) result.Add(change.Value);
            result.Sort((left, right) => left.Tag.Directory == right.Tag.Directory ? left.Tag.Id.CompareTo(right.Tag.Id) : left.Tag.Directory.CompareTo(right.Tag.Directory));
            return result.AsReadOnly();
        }
    }

    /// <summary>Returns a typed field, or null when absent.</summary>
    public OfficeExifValue? GetExifValue(OfficeExifTag tag) {
        foreach (OfficeExifValue field in ExifValues) if (field.Tag.Equals(tag)) return field;
        return null;
    }
    /// <summary>Sets an Exif scalar, text, rational, or typed array using the tag's representation.</summary>
    public void SetExifValue(OfficeExifTag tag, object value) => SetExifValue(tag, value, CancellationToken.None);
    private void SetExifValue(OfficeExifTag tag, object value, CancellationToken token) {
        if (IsStructural(tag.Id)) throw new ArgumentException("Exif structural pointers cannot be edited as ordinary fields.", nameof(tag));
        byte[] encoded = OfficeExifProfileCodec.EncodeValue(tag.DataType, value, true, out uint count, token);
        object decoded = OfficeExifProfileCodec.Decode(encoded, 0, checked((int)count), tag.DataType, true, token);
        var change = new OfficeExifValue(tag, decoded);
        if (tag.Equals(OfficeExifTag.XResolution) || tag.Equals(OfficeExifTag.YResolution) || tag.Equals(OfficeExifTag.ResolutionUnit)) {
            SetResolutionAndDensityFields(ResolveExifDensityEdit(tag, decoded), restoreRemoved: false, edited: change);
        } else {
            _changes[tag] = change;
        }
        _removed.Remove(tag);
        _tiffOpaqueOffsets.Remove(tag);
    }
    /// <summary>Removes a field and returns whether it was present.</summary>
    public bool RemoveExifValue(OfficeExifTag tag) {
        if (IsStructural(tag.Id)) throw new ArgumentException("Exif structural pointers cannot be edited as ordinary fields.", nameof(tag));
        bool present = GetExifValue(tag) != null;
        _changes.Remove(tag); _removed.Add(tag); _tiffOpaqueOffsets.Remove(tag); return present;
    }
    /// <summary>Removes the entire Exif profile, including opaque and thumbnail data.</summary>
    /// <remarks>When subsequently applied to a TIFF, its original rendering orientation is retained unless explicitly removed with <see cref="RemoveExifValue"/> after clearing the profile.</remarks>
    public void ClearExif() { _exif = null; _changes.Clear(); _removed.Clear(); _tiffOpaqueOffsets.Clear(); }
    /// <summary>Returns a classic TIFF profile containing the current edits, or null when Exif is absent.</summary>
    /// <remarks>Unedited data retains its original offsets. Edited values are erased from obsolete storage when ranges are exclusive; overlapping values are rejected.</remarks>
    public byte[]? EncodeExifProfile() => EncodeExifProfile(CancellationToken.None);
    /// <summary>Encodes classic TIFF Exif metadata while observing cancellation during directory and value processing.</summary>
    public byte[]? EncodeExifProfile(CancellationToken cancellationToken) => EncodeExifProfile(cancellationToken, 0L);
    internal byte[]? EncodeExifProfile(CancellationToken cancellationToken, long additionallyRetainedBytes) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_tiffOpaqueOffsets.Count != 0) throw new NotSupportedException("TIFF-relative maker-note offsets cannot be exported to a different container. Retain the original TIFF for lossless edits, or explicitly remove or replace the maker-note field.");
        return _exif == null && _changes.Count == 0 ? null : OfficeExifProfileCodec.Encode(_exif, _changes, _removed, cancellationToken: cancellationToken, additionallyRetainedBytes: additionallyRetainedBytes);
    }
    /// <summary>Reads a classic TIFF Exif profile or a JPEG-prefixed Exif profile.</summary>
    public static OfficeImageMetadata ParseExifProfile(byte[] profile) => new OfficeImageMetadata { ExifProfile = profile };
    /// <summary>Reads classic TIFF or JPEG-prefixed Exif metadata while observing cancellation.</summary>
    public static OfficeImageMetadata ParseExifProfile(byte[] profile, CancellationToken cancellationToken) {
        var metadata = new OfficeImageMetadata();
        metadata.SetExifProfile(profile, cancellationToken);
        return metadata;
    }
    internal static OfficeImageMetadata ParseExifProfile(byte[] profile, long additionallyRetainedBytes, CancellationToken cancellationToken) {
        var metadata = new OfficeImageMetadata { _exif = OfficeExifProfileCodec.Parse(profile, cancellationToken: cancellationToken, additionallyRetainedBytes: additionallyRetainedBytes) };
        // The caller's container is retained during parsing, not by this independently owned snapshot.
        metadata._exif.RetainedManagedBytes -= additionallyRetainedBytes;
        ReadExifResolution(metadata);
        return metadata;
    }
    private void SetExifProfile(byte[]? profile, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var parsed = new OfficeImageMetadata { _exif = profile == null || profile.Length == 0 ? null : OfficeExifProfileCodec.Parse(profile, cancellationToken: token) };
        ReadExifResolution(parsed);
        token.ThrowIfCancellationRequested();
        _exif = parsed._exif;
        _changes.Clear(); _removed.Clear(); _tiffOpaqueOffsets.Clear();
        _resolution = parsed._resolution;
    }

    /// <summary>Creates an independent copy of profiles, typed fields, pending edits, and TIFF preservation constraints.</summary>
    public OfficeImageMetadata Clone() {
        var clone = new OfficeImageMetadata { _exif = _exif, _xmp = Copy(_xmp), _icc = Copy(_icc), _iptc = Copy(_iptc), _resolution = _resolution };
        foreach (KeyValuePair<OfficeExifTag, OfficeExifValue> change in _changes) clone._changes.Add(change.Key, change.Value);
        foreach (OfficeExifTag removed in _removed) clone._removed.Add(removed);
        foreach (OfficeExifTag opaque in _tiffOpaqueOffsets) clone._tiffOpaqueOffsets.Add(opaque);
        return clone;
    }

    /// <summary>Reads image profiles and physical resolution from a supported encoded image.</summary>
    public static OfficeImageMetadata Read(byte[] encodedBytes, CancellationToken cancellationToken = default) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(encodedBytes.Length)) throw new FormatException("Image bytes exceed the metadata-read limit.");
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info)) throw new FormatException("The image container is malformed or unsupported.");
        OfficeImageMetadataSnapshot snapshot = OfficeImageMetadataInspector.Inspect(encodedBytes, info.Format, 0L, cancellationToken);
        var metadata = new OfficeImageMetadata();
        if (snapshot.Exif != null) metadata.SetExifProfile(snapshot.Exif, cancellationToken);
        metadata.IccProfile = snapshot.Icc;
        if (snapshot.Xmp != null) {
            int skip = info.Format == OfficeImageFormat.Jpeg && StartsWith(snapshot.Xmp, JpegXmpPrefix) ? JpegXmpPrefix.Length : 0;
            var xmp = new byte[snapshot.Xmp.Length - skip]; Buffer.BlockCopy(snapshot.Xmp, skip, xmp, 0, xmp.Length); metadata.XmpProfile = xmp;
        }
        if (snapshot.PhysicalDpiX.HasValue && snapshot.PhysicalDpiY.HasValue) metadata._resolution = new OfficeImageResolution(snapshot.PhysicalDpiX.Value, snapshot.PhysicalDpiY.Value);
        if (info.Format == OfficeImageFormat.Jpeg) metadata.IptcProfile = ReadJpegIptc(encodedBytes, cancellationToken);
        if (info.Format == OfficeImageFormat.Png) ReadPngProfiles(encodedBytes, metadata, cancellationToken);
        if (info.Format == OfficeImageFormat.Webp) ReadWebpProfiles(encodedBytes, metadata, cancellationToken);
        if (info.Format == OfficeImageFormat.Tiff) ReadTiffProfiles(encodedBytes, metadata, cancellationToken);
        if (info.Format == OfficeImageFormat.Bmp) ReadBmpProfiles(encodedBytes, metadata, cancellationToken);
        if (info.Format == OfficeImageFormat.Gif) ReadGifProfiles(encodedBytes, metadata, cancellationToken);
        ReadNativeResolution(encodedBytes, info.Format, metadata, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return metadata;
    }

    /// <summary>Replaces supported primary-image profiles and density without recompressing image pixels.</summary>
    /// <remarks>Supports JPEG, PNG, WebP, and TIFF profile families, GIF XMP/ICC, and BMP ICC/density. TIFF edits preserve page links and encoding fields; TIFF-relative maker notes require their original container. Container-specific metadata outside these profile families is preserved. Extended XMP is replaced by the supplied standard XMP packet.</remarks>
    public static byte[] Apply(byte[] encodedBytes, OfficeImageMetadata metadata, CancellationToken cancellationToken = default) => ApplyCore(encodedBytes, metadata, cancellationToken, 0L);

    internal static byte[] ApplyCore(byte[] encodedBytes, OfficeImageMetadata metadata, CancellationToken cancellationToken, long additionallyRetainedBytes, bool writeDensity = true) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        if (metadata == null) throw new ArgumentNullException(nameof(metadata));
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(encodedBytes.Length)) throw new FormatException("Image bytes exceed the metadata-edit limit.");
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info)) throw new FormatException("The image container is malformed or unsupported.");
        metadata.ValidateResolution();
        if (info.Format != OfficeImageFormat.Jpeg && info.Format != OfficeImageFormat.Png && info.Format != OfficeImageFormat.Webp && info.Format != OfficeImageFormat.Tiff && info.Format != OfficeImageFormat.Bmp && info.Format != OfficeImageFormat.Gif) throw new NotSupportedException("The image container does not support lossless replacement of these profile families.");
        if (metadata._icc != null && !OfficeIccProfileValidator.TryValidate(metadata._icc, 0, metadata._icc.Length, cancellationToken)) throw new FormatException("The ICC profile is malformed.");
        byte[] result = info.Format switch {
            OfficeImageFormat.Jpeg => RewriteJpeg(encodedBytes, metadata, cancellationToken, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _, additionallyRetainedBytes, writeDensity),
            OfficeImageFormat.Png => RewritePng(encodedBytes, metadata, cancellationToken, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _, additionallyRetainedBytes, writeDensity),
            OfficeImageFormat.Webp => RewriteWebp(encodedBytes, metadata, cancellationToken, OfficeImageMetadataProfileKinds.All & ~OfficeImageMetadataProfileKinds.C2pa, out _, additionallyRetainedBytes),
            OfficeImageFormat.Tiff => RewriteTiff(encodedBytes, metadata, cancellationToken, preserveDensity: writeDensity, additionallyRetainedBytes: additionallyRetainedBytes),
            OfficeImageFormat.Bmp => RewriteBmp(encodedBytes, metadata, OfficeImageMetadataProfileKinds.All, cancellationToken, out _, additionallyRetainedBytes, writeDensity),
            OfficeImageFormat.Gif => RewriteGif(encodedBytes, metadata, OfficeImageMetadataProfileKinds.All, cancellationToken, out _, additionallyRetainedBytes, writeDensity),
            _ => throw new NotSupportedException("The image container does not support lossless replacement of these profile families.")
        };
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(result.Length)) throw new FormatException("Edited image exceeds the encoded-size limit.");
        return result;
    }
    private void ValidateResolution() {
        if (HorizontalResolution <= 0 || VerticalResolution <= 0 || double.IsNaN(HorizontalResolution) || double.IsNaN(VerticalResolution) || double.IsInfinity(HorizontalResolution) || double.IsInfinity(VerticalResolution) || ResolutionUnits < OfficeImageResolutionUnit.AspectRatio || ResolutionUnits > OfficeImageResolutionUnit.PixelsPerMeter) throw new ArgumentOutOfRangeException(nameof(HorizontalResolution), "Image resolution must be finite, positive, and use a defined unit.");
    }
    private double? ToPhysicalDpi(double resolution) {
        ValidateResolution();
        return ResolutionUnits switch {
            OfficeImageResolutionUnit.AspectRatio => null,
            OfficeImageResolutionUnit.PixelsPerInch => resolution,
            OfficeImageResolutionUnit.PixelsPerCentimeter => resolution * 2.54D,
            OfficeImageResolutionUnit.PixelsPerMeter => resolution * 0.0254D,
            _ => throw new ArgumentOutOfRangeException(nameof(ResolutionUnits))
        };
    }
    private static bool IsStructural(ushort id) => id == 34665 || id == 34853 || id == 40965 || id == 330 || id == 513 || id == 514;
    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
    private static byte[]? BoundedCopy(byte[]? value) { if (value != null && value.Length > OfficeExifProfileCodec.MaximumProfileBytes) throw new ArgumentOutOfRangeException(nameof(value), "The metadata profile exceeds the size limit."); return value == null || value.Length == 0 ? null : (byte[])value.Clone(); }
}
