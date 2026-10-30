using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Accounts for semantic differences before exact copying or preparing bounded native source edits.</summary>
internal sealed partial class PublisherNativeWriter {
    private readonly PublisherPublicationState _current, _baseline;
    private readonly PublisherImportedResidue _imported;
    private readonly IReadOnlyDictionary<uint, PublisherEscherShape> _artwork;
    private readonly CancellationToken _token;
    private readonly List<OfficeConversionFidelityDiagnostic> _diagnostics = new();
    private readonly List<PublisherQuillStoryStyles> _styleDependencies = new();
    private readonly Dictionary<int, (uint Original, uint Replacement)> _patches = new();
    private byte[]? _quillBytes;
    private readonly OfficeLegacyImportLimits _sourceLimits;
    private readonly long _maximumBytes;

    private PublisherNativeWriter(PublisherPublicationState current, PublisherImportedResidue imported,
        OfficeLegacyImportLimits sourceLimits, long maximumBytes, CancellationToken token) {
        _current = current; _baseline = imported.Baseline!; _imported = imported; _token = token;
        _sourceLimits = sourceLimits; _maximumBytes = maximumBytes;
        _artwork = imported.Artwork.Shapes.ToDictionary(shape => shape.Id);
    }

    internal static PublisherNativePreparation Prepare(PublisherPublicationState current, PublisherImportedResidue? imported,
        OfficeLegacyImportLimits sourceLimits, long maximumBytes, CancellationToken token) {
        if (maximumBytes < 1 || maximumBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        token.ThrowIfCancellationRequested();
        if (imported?.NativeSource == null || imported.Baseline == null)
            return new PublisherNativePreparation(null, new[] { Finding("PUB_NATIVE_SOURCE_REQUIRED",
                "Native preparation requires an imported source and its semantic baseline. Source-free authoring is not a native writer.", OfficeConversionLossKind.Failure, "Contents") });
        var writer = new PublisherNativeWriter(current, imported, sourceLimits, maximumBytes, token);
        writer.Compare();
        if (writer._diagnostics.Count != 0) return new PublisherNativePreparation(null, writer._diagnostics, writer._styleDependencies);
        PublisherNativeSource source = imported.NativeSource;
        if (writer._patches.Count == 0 && writer._quillBytes == null) {
            if (source.ByteLength > maximumBytes) throw new InvalidDataException("Publisher native output byte limit exceeded.");
            token.ThrowIfCancellationRequested();
            return new PublisherNativePreparation(source.CopyBytes(), new[] { Finding("PUB_NATIVE_EXACT_COPY",
                "The semantic model is unchanged. Original native bytes are copied exactly; native validity and signatures are not reassessed.", OfficeConversionLossKind.None, "Contents") });
        }
        if (!source.CanRewrite(sourceLimits, token)) {
            writer.Reject("Contents", "Edited output is unavailable for an invalid directory tree, nonempty active-content carriers, signatures or protected storage.");
            return new PublisherNativePreparation(null, writer._diagnostics, writer._styleDependencies);
        }
        var streams = new Dictionary<string, byte[]>();
        byte[] bytes = writer._patches.Count == 0 ? Array.Empty<byte>() : (byte[])source.EscherBytes.Clone();
        var data = new PublisherBinaryData(bytes, "Escher/EscherStm");
        foreach (var patch in writer._patches) {
            token.ThrowIfCancellationRequested();
            if (data.U32(patch.Key) != patch.Value.Original) {
                writer.Reject("Escher/EscherStm", "The retained native source no longer matches its guarded patch slots.");
                return new PublisherNativePreparation(null, writer._diagnostics, writer._styleDependencies);
            }
            uint value = patch.Value.Replacement;
            for (int index = 0; index < 4; index++) bytes[patch.Key + index] = (byte)(value >> (index * 8));
        }
        if (writer._patches.Count != 0) streams.Add("Escher/EscherStm", bytes);
        if (writer._quillBytes != null) streams.Add("Quill/QuillSub/CONTENTS", writer._quillBytes);
        byte[] output = source.Rewrite(streams, maximumBytes, token);
        token.ThrowIfCancellationRequested();
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        if (writer._patches.Count != 0) diagnostics.Add(Finding("PUB_NATIVE_FIXED_PATCH_UNQUALIFIED",
            "Guarded fixed-size OfficeArt slots changed. Streams outside the prepared replacements and directory metadata are retained. Repair-free opening, native edit acceptance and opaque semantic preservation require independent qualification.",
            OfficeConversionLossKind.Unassessed, "Escher/EscherStm"));
        if (writer._quillBytes != null) diagnostics.Add(Finding("PUB_NATIVE_QUILL_TEXT_UNQUALIFIED",
            "The bounded single-story Quill profile was rewritten with unchanged paragraph topology, retained source OPLs and supported size/bold edits. Other logical streams, including cached previews, are retained. Native opening, editing, resave/reopen and opaque record preservation remain unassessed.",
            OfficeConversionLossKind.Unassessed, "Quill"));
        return new PublisherNativePreparation(output, diagnostics, writer._styleDependencies);
    }

    private static OfficeConversionFidelityDiagnostic Finding(string code, string message, OfficeConversionLossKind kind, string location) =>
        new(code, message, kind, "OfficeIMO.Publisher", location);
    private void Reject(string location, string message) =>
        _diagnostics.Add(Finding("PUB_NATIVE_EDIT_UNSUPPORTED", message, OfficeConversionLossKind.Failure, location));

    private PublisherNativeShapeBinding? Binding(uint id) {
        PublisherNativeShapeBinding? binding = _artwork.TryGetValue(id, out PublisherEscherShape? shape) ? shape.NativeBinding : null;
        if (binding == null || !binding.Matches(new PublisherBinaryData(_imported.NativeSource!.EscherBytes, "Escher/EscherStm"))) {
            Reject("Contents/object/" + id, "The object has no matching exact native shape binding."); return null;
        }
        return binding;
    }

    private void Patch(PublisherNativeValueSlot slot, uint value) {
        if (slot.Original == value) return;
        if (_patches.TryGetValue(slot.Offset, out var prior) && (prior.Original != slot.Original || prior.Replacement != value)) {
            Reject("Escher/EscherStm", "Current semantic owners disagree about one native slot."); return;
        }
        _patches[slot.Offset] = (slot.Original, value);
    }

    private void Bounds(uint id, uint? owner, double x, double y, double width, double height) {
        _token.ThrowIfCancellationRequested();
        PublisherNativeShapeBinding? binding = Binding(id);
        PublisherPageState? page = owner.HasValue ? _baseline.Pages.SingleOrDefault(item => item.Id == owner.Value) : null;
        if (binding == null) return;
        if (page == null || !page.SupportsPlacementEditing || binding.Anchor.Count != 4
            || !Emu(x, page.RegistrationX, out int left) || !Emu(y, page.RegistrationY, out int top)
            || !Emu(x + width, page.RegistrationX, out int right) || !Emu(y + height, page.RegistrationY, out int bottom)) {
            Reject("Contents/object/" + id, "The current bounds have no qualified complete anchor or exact signed-EMU representation."); return;
        }
        // Native endpoint direction is independent of the semantic bounding rectangle.
        if (unchecked((int)binding.Anchor[0].Original) > unchecked((int)binding.Anchor[2].Original)) (left, right) = (right, left);
        if (unchecked((int)binding.Anchor[1].Original) > unchecked((int)binding.Anchor[3].Original)) (top, bottom) = (bottom, top);
        uint[] values = { unchecked((uint)left), unchecked((uint)top), unchecked((uint)right), unchecked((uint)bottom) };
        for (int index = 0; index < values.Length; index++) Patch(binding.Anchor[index], values[index]);
    }

    private static bool Emu(double point, double origin, out int result) {
        double scaled = (point - origin) * 12700, rounded = Math.Round(scaled);
        result = 0;
        if (double.IsNaN(scaled) || double.IsInfinity(scaled) || rounded < int.MinValue || rounded > int.MaxValue
            || Math.Abs(scaled - rounded) > Math.Max(0.0000001, Math.Abs(scaled) * 0.000000000000002)) return false;
        result = (int)rounded; return true;
    }

    private void Paint(uint id, OfficeShape before, OfficeShape after) {
        PublisherNativeShapeBinding? binding = Binding(id); if (binding == null) return;
        void Scalar(ushort property, uint value) {
            PublisherNativePropertySlot[] slots = binding.Properties.Where(item => item.Id == property).ToArray();
            if (slots.Length != 1 || (slots[0].RawOperationId & 0xC000) != 0) {
                Reject("Contents/object/" + id, $"Solid property 0x{property:X4} has no unambiguous existing local scalar slot."); return;
            }
            Patch(slots[0].Value, value);
        }
        void Color(OfficeColor? prior, OfficeColor? value, ushort property) {
            if (prior == value) return;
            // Visibility changes require a separately qualified boolean-mask writer.
            if (!prior.HasValue || !value.HasValue || value.Value.A != 255) {
                Reject("Contents/object/" + id, "Paint visibility or alpha changes are outside this fixed scalar slice."); return;
            }
            Scalar(property, (uint)(value.Value.R | value.Value.G << 8 | value.Value.B << 16));
        }
        void Fixed(double prior, double value, ushort property, double scale) {
            if (prior == value) return;
            double scaled = value * scale;
            if (scaled < 0 || scaled > int.MaxValue || scaled != Math.Round(scaled)) {
                Reject("Contents/object/" + id, "The solid paint value has no exact native scalar representation."); return;
            }
            Scalar(property, (uint)scaled);
        }
        Color(before.FillColor, after.FillColor, 0x0181);
        Color(before.StrokeColor, after.StrokeColor, 0x01C0);
        Fixed(before.FillOpacity ?? 1, after.FillOpacity ?? 1, 0x0182, 65536);
        Fixed(before.StrokeOpacity ?? 1, after.StrokeOpacity ?? 1, 0x01C1, 65536);
        Fixed(before.StrokeWidth, after.StrokeWidth, 0x01CB, 12700);
    }
}
