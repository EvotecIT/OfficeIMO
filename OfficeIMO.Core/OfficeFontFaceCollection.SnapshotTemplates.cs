using System;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private static readonly ConditionalWeakTable<byte[], SnapshotTemplateSlot> SnapshotTemplates = new();

    /// <summary>
    /// Retains an already private, immutable font snapshot for shared layout. Only the accepted
    /// native static program is reused; each operation pays the ordinary logical decoded-byte
    /// charge and owns its face registration. Callers must never pass a mutable source buffer.
    /// </summary>
    internal bool TryAddImmutableSnapshotBounded(string? familyName, byte[]? data,
        OfficeFontFaceDescriptor descriptor, int maximumDecodedBytes,
        out int decodedBytes, out string? error) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumDecodedBytes);
        if (FontProgramProvider != null || FontVariationResolver != null
            || string.IsNullOrWhiteSpace(familyName) || data == null || data.Length == 0
            || OfficeFontContainerDecoder.Detect(data) != OfficeFontContainerFormat.OpenType
            || HasTrueTypeCollectionSignature(data)) {
            return TryAddBounded(familyName, data, descriptor, OfficeFontUnicodeRangeSet.All,
                maximumDecodedBytes, out decodedBytes, out error);
        }

        SnapshotTemplateSlot slot = SnapshotTemplates.GetValue(data, static _ => new SnapshotTemplateSlot());
        lock (slot) {
            if (slot.Template != null && slot.Template.DecodedBytes <= maximumDecodedBytes) {
                RegisterTemplate(familyName!.Trim(), slot.Template, descriptor);
                decodedBytes = slot.Template.DecodedBytes;
                error = null;
                return true;
            }

            // Failed admission uses the original decoder/parser, preserving its limit checks,
            // error ordering and messages. A small first budget never poisons a later request.
            if (!TryAddBounded(familyName, data, descriptor, OfficeFontUnicodeRangeSet.All,
                    maximumDecodedBytes, out decodedBytes, out error)) return false;
            if (slot.Template != null || slot.Declined) return true;

            string normalizedFamily = familyName!.Trim();
            string resourceFamily = descriptor == OfficeFontFaceDescriptor.Regular
                ? normalizedFamily : CreateResourceFamilyName(normalizedFamily, descriptor, OfficeFontUnicodeRangeSet.All);
            OfficeFontFace face = _faces.FindLast(candidate => candidate.FamilyName == normalizedFamily
                && candidate.ResourceFamilyName == resourceFamily && candidate.Descriptor == descriptor)!;
            bool staticNative = face.CanEmbedAsStaticPdfFont &&
                (face.ParsedFont is OfficeTrueTypeFont trueType && !trueType.IsVariableFont
                    || face.ParsedFont is OfficeOpenTypeCffFont cff && !cff.IsVariable);
            OfficeOpenTypeReader? reader = staticNative ? OfficeOpenTypeReader.TryCreate(face.DataSnapshot) : null;
            if (reader != null && !reader.TryGetTable("trak", out _, out _)) {
                // The wrapper is private to the cache; its immutable payload is shared with
                // fresh operation wrappers. No caller array or collection is retained.
                slot.Template = new FaceTemplate(face.Clone(), decodedBytes);
            } else {
                slot.Declined = true;
            }
            return true;
        }
    }

    private sealed class SnapshotTemplateSlot {
        internal FaceTemplate? Template { get; set; }
        internal bool Declined { get; set; }
    }

    private sealed class FaceTemplate {
        internal FaceTemplate(OfficeFontFace face, int decodedBytes) {
            Face = face;
            DecodedBytes = decodedBytes;
        }

        internal OfficeFontFace Face { get; }
        internal int DecodedBytes { get; }
    }
}
