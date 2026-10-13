using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Finite input and retained-evidence limits for passive image metadata inspection.</summary>
public sealed class OfficeImageMetadataInspectionOptions {
    /// <summary>Maximum encoded input bytes. Default: 16 MiB; hard ceiling: 128 MiB.</summary>
    public int MaximumInputBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum bytes in one extracted profile. Default: 4 MiB; hard ceiling: 16 MiB.</summary>
    public int MaximumProfileBytes { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum retained Exif field descriptors across the inspected profiles. Default: 512.</summary>
    public int MaximumExifFields { get; set; } = 512;
    /// <summary>Maximum bytes decoded from one Exif value. Larger values retain their descriptor only.</summary>
    public int MaximumExifValueBytes { get; set; } = 64 * 1024;
    /// <summary>Maximum aggregate encoded bytes of retained Exif values. Default: 4 MiB; hard ceiling: 64 MiB.</summary>
    public int MaximumRetainedExifBytes { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum retained XMP values. Default: 512.</summary>
    public int MaximumXmpProperties { get; set; } = 512;
    /// <summary>Maximum XML nodes examined before materialization. Default: 4096.</summary>
    public int MaximumXmlNodes { get; set; } = 4096;
    /// <summary>Maximum characters retained in one XMP value. Default: 4096.</summary>
    public int MaximumValueCharacters { get; set; } = 4096;
    /// <summary>Maximum aggregate retained XMP value and language characters. Default: one million.</summary>
    public int MaximumRetainedXmpCharacters { get; set; } = 1_000_000;
    /// <summary>Maximum retained ICC tag descriptors. All tag ranges are still validated.</summary>
    public int MaximumIccTags { get; set; } = 256;

    /// <summary>Returns an independently owned validated policy snapshot.</summary>
    public OfficeImageMetadataInspectionOptions Clone() => Snapshot();

    internal OfficeImageMetadataInspectionOptions Snapshot() {
        if (MaximumInputBytes < 1 || MaximumInputBytes > OfficeRasterGuards.MaximumEncodedBytes ||
            MaximumProfileBytes < 1 || MaximumProfileBytes > OfficeExifProfileCodec.MaximumProfileBytes ||
            MaximumExifFields < 1 || MaximumExifFields > 65535 || MaximumExifValueBytes < 1 || MaximumExifValueBytes > MaximumProfileBytes ||
            MaximumRetainedExifBytes < 1 || MaximumRetainedExifBytes > 64 * 1024 * 1024 ||
            MaximumXmpProperties < 1 || MaximumXmpProperties > 65535 || MaximumXmlNodes < 1 || MaximumXmlNodes > 65535 ||
            MaximumValueCharacters < 1 || MaximumValueCharacters > 1024 * 1024 || MaximumRetainedXmpCharacters < 1 || MaximumRetainedXmpCharacters > 4 * 1024 * 1024 || MaximumIccTags < 1 || MaximumIccTags > 65535) {
            throw new ArgumentOutOfRangeException(nameof(OfficeImageMetadataInspectionOptions), "Metadata inspection requires finite positive limits within the documented ceilings.");
        }
        return (OfficeImageMetadataInspectionOptions)MemberwiseClone();
    }
}

/// <summary>A source Exif field, including duplicate identities and descriptors of omitted values.</summary>
public sealed class OfficeExifFieldEvidence {
    internal OfficeExifFieldEvidence(OfficeExifDirectory directory, ushort id, ushort dataType, uint count,
        int directoryOffset, int entryOffset, int valueOffset, int valueLength, string source, OfficeExifValue? value, string? omission) {
        Directory = directory; TagId = id; RawDataType = dataType; Count = count;
        DirectoryOffset = directoryOffset; EntryOffset = entryOffset; ValueOffset = valueOffset;
        ValueLengthBytes = valueLength; Source = source; Value = value; Omission = omission;
    }
    /// <summary>Semantic directory reached by the classic Exif profile.</summary>
    public OfficeExifDirectory Directory { get; }
    /// <summary>Original unsigned tag identifier.</summary>
    public ushort TagId { get; }
    /// <summary>Original TIFF type code, including unsupported type 13.</summary>
    public ushort RawDataType { get; }
    /// <summary>Original declared item count, including zero counts.</summary>
    public uint Count { get; }
    /// <summary>Directory offset relative to the TIFF profile, excluding JPEG Exif framing.</summary>
    public int DirectoryOffset { get; }
    /// <summary>Field-entry offset relative to the TIFF profile.</summary>
    public int EntryOffset { get; }
    /// <summary>Value offset relative to the TIFF profile.</summary>
    public int ValueOffset { get; }
    /// <summary>Declared encoded value length.</summary>
    public int ValueLengthBytes { get; }
    /// <summary>Carrier identity; profile-relative offsets are not file offsets.</summary>
    public string Source { get; }
    /// <summary>Typed value, or null when unsupported, structural, or omitted by policy.</summary>
    public OfficeExifValue? Value { get; }
    /// <summary>Stable reason why the value was not decoded, or null.</summary>
    public string? Omission { get; }
}

/// <summary>Authored image density distinguished from a rendering default.</summary>
public sealed class OfficeImageResolutionEvidence {
    internal OfficeImageResolutionEvidence(bool declared, string source, OfficeImageResolution? resolution, bool valid, bool presenceKnown = true) {
        IsDeclared = declared; Source = source; Resolution = resolution; IsValid = valid; IsPresenceKnown = presenceKnown;
    }
    /// <summary>Whether a density or aspect-ratio carrier was found.</summary>
    public bool IsDeclared { get; }
    /// <summary>Whether the supported primary carriers could establish presence or absence; false when inspection was limited.</summary>
    public bool IsPresenceKnown { get; }
    /// <summary>Carrier used for the effective value, or "none".</summary>
    public string Source { get; }
    /// <summary>Whether the carrier contains complete positive supported values.</summary>
    public bool IsValid { get; }
    /// <summary>Validated authored value; null when missing or invalid. No 96-DPI default is supplied.</summary>
    public OfficeImageResolution? Resolution { get; }
}

/// <summary>Passive metadata evidence for a bounded image; it does not establish pixel integrity or authorship.</summary>
public sealed class OfficeImageMetadataInspection {
    internal OfficeImageMetadataInspection(OfficeImageFormat format, List<OfficeExifFieldEvidence> fields,
        List<OfficeExifTimestamp> timestamps, OfficeExifGpsCoordinate? gps, OfficeImageResolutionEvidence resolution,
        List<OfficeXmpProperty> xmp, OfficeIccProfileInspection? icc, List<string> issues) {
        Format = format; ExifFields = fields.AsReadOnly(); Timestamps = timestamps.AsReadOnly(); GpsCoordinate = gps;
        Resolution = resolution; XmpProperties = xmp.AsReadOnly(); IccProfile = icc; Issues = issues.AsReadOnly();
    }
    /// <summary>Identified image container.</summary>
    public OfficeImageFormat Format { get; }
    /// <summary>Retained source fields, without silently removing duplicate tags.</summary>
    public IReadOnlyList<OfficeExifFieldEvidence> ExifFields { get; }
    /// <summary>Normalized timestamp evidence, retaining unknown timezone and malformed source values.</summary>
    public IReadOnlyList<OfficeExifTimestamp> Timestamps { get; }
    /// <summary>Validated geographic coordinates, or null when incomplete or contradictory.</summary>
    public OfficeExifGpsCoordinate? GpsCoordinate { get; }
    /// <summary>Authored resolution evidence.</summary>
    public OfficeImageResolutionEvidence Resolution { get; }
    /// <summary>Bounded scalar, array-item and language-alternative XMP properties.</summary>
    public IReadOnlyList<OfficeXmpProperty> XmpProperties { get; }
    /// <summary>Embedded ICC structural evidence, independently of transform support.</summary>
    public OfficeIccProfileInspection? IccProfile { get; }
    /// <summary>Stable corruption, ambiguity, omission and profile-scope diagnostics.</summary>
    public IReadOnlyList<string> Issues { get; }
    /// <summary>Whether all evidence in the documented profile was retained without an issue.</summary>
    public bool IsComplete => Issues.Count == 0;
}
