using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One validated ICC tag-table entry.</summary>
public sealed class OfficeIccTagEvidence {
    internal OfficeIccTagEvidence(string signature, uint offset, uint length, string type) { Signature = signature; Offset = offset; LengthBytes = length; TypeSignature = type; }
    /// <summary>Four-byte tag signature.</summary>
    public string Signature { get; }
    /// <summary>Profile-relative offset.</summary>
    public uint Offset { get; }
    /// <summary>Declared tag length.</summary>
    public uint LengthBytes { get; }
    /// <summary>Four-byte encoded data-type signature.</summary>
    public string TypeSignature { get; }
}

/// <summary>ICC structural evidence; conversion support and a valid structural envelope are separate facts.</summary>
public sealed class OfficeIccProfileInspection {
    internal OfficeIccProfileInspection(bool valid, bool? supported, uint declaredLength, byte major, byte minor,
        string profileClass, string colorSpace, string pcs, uint intent, uint tagCount, List<OfficeIccTagEvidence> tags, List<string> issues) {
        IsStructurallyValid = valid; HasSupportedInputTransform = supported; DeclaredLengthBytes = declaredLength;
        MajorVersion = major; MinorVersion = minor; ProfileClassSignature = profileClass; DeviceColorSpaceSignature = colorSpace;
        ConnectionSpaceSignature = pcs; RenderingIntent = intent; DeclaredTagCount = tagCount; Tags = tags.AsReadOnly(); Issues = issues.AsReadOnly();
    }
    /// <summary>Whether the canonical ICC size/signature/reserved-field/tag-range validation passed.</summary>
    public bool IsStructurallyValid { get; }
    /// <summary>Whether the current managed conversion engine accepts an input transform; null when structural validation failed or was limited.</summary>
    public bool? HasSupportedInputTransform { get; }
    /// <summary>Header-declared profile length, when the header is available.</summary>
    public uint DeclaredLengthBytes { get; }
    /// <summary>Header-declared major version.</summary>
    public byte MajorVersion { get; }
    /// <summary>Header-declared minor version.</summary>
    public byte MinorVersion { get; }
    /// <summary>Raw four-byte device-class signature.</summary>
    public string ProfileClassSignature { get; }
    /// <summary>Raw four-byte device-color-space signature.</summary>
    public string DeviceColorSpaceSignature { get; }
    /// <summary>Raw four-byte profile connection space.</summary>
    public string ConnectionSpaceSignature { get; }
    /// <summary>Header-declared rendering intent.</summary>
    public uint RenderingIntent { get; }
    /// <summary>Original tag count, independently of retention limits.</summary>
    public uint DeclaredTagCount { get; }
    /// <summary>Retained validated tag ranges.</summary>
    public IReadOnlyList<OfficeIccTagEvidence> Tags { get; }
    /// <summary>Stable validity, unsupported-transform and retention diagnostics.</summary>
    public IReadOnlyList<string> Issues { get; }
}

/// <summary>Inspects embedded or standalone ICC bytes without applying a color transformation.</summary>
public static class OfficeIccProfileInspector {
    /// <summary>Returns bounded header and tag-table evidence using the shared ICC validator and conversion parser.</summary>
    public static OfficeIccProfileInspection Inspect(byte[] profile, OfficeImageMetadataInspectionOptions? options = null, CancellationToken cancellationToken = default) {
        if (profile == null) throw new ArgumentNullException(nameof(profile));
        OfficeImageMetadataInspectionOptions policy = (options ?? new OfficeImageMetadataInspectionOptions()).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        var tags = new List<OfficeIccTagEvidence>(); var issues = new List<string>();
        bool bounded = profile.Length <= policy.MaximumProfileBytes;
        bool valid = bounded && OfficeIccProfileValidator.TryValidate(profile, 0, profile.Length, cancellationToken);
        bool? supported = valid ? OfficeIccColorProfile.TryCreate(profile, out _) : null;
        cancellationToken.ThrowIfCancellationRequested();
        if (!bounded) issues.Add("icc:profile-byte-limit");
        else if (!valid) issues.Add("icc:invalid-structure");
        else if (supported == false) issues.Add("icc:input-transform-unsupported");
        uint count = profile.Length >= 132 ? Read(profile, 128) : 0;
        if (valid) for (int i = 0; i < count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (tags.Count >= policy.MaximumIccTags) { issues.Add("icc:tag-retention-limit"); break; }
            int at = 132 + i * 12; uint offset = Read(profile, at + 4), length = Read(profile, at + 8);
            tags.Add(new OfficeIccTagEvidence(Signature(profile, at), offset, length, Signature(profile, (int)offset)));
        }
        return new OfficeIccProfileInspection(valid, supported, profile.Length >= 4 ? Read(profile, 0) : 0,
            profile.Length >= 12 ? profile[8] : (byte)0, profile.Length >= 12 ? (byte)(profile[9] >> 4) : (byte)0,
            Signature(profile, 12), Signature(profile, 16), Signature(profile, 20), profile.Length >= 68 ? Read(profile, 64) : 0, count, tags, issues);
    }
    private static uint Read(byte[] bytes, int offset) => (uint)OfficeExifProfileCodec.Read(bytes, offset, 4, false);
    private static string Signature(byte[] bytes, int offset) => offset >= 0 && offset <= bytes.Length - 4 ? Encoding.ASCII.GetString(bytes, offset, 4) : string.Empty;
}
