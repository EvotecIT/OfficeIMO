using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    /// <summary>Owns source analysis, feature selection and work limits for one complete script-shaping request.</summary>
    internal sealed class ScriptShapingContext {
        private readonly OfficeOpenTypeSubstitution _owner;
        private readonly Dictionary<string, List<ArabicFeatureLookup>?> _lookups = new(StringComparer.Ordinal);
        private string?[]? _joining;
        private readonly int[]? _joiningControls;
        internal int Operations, Inspections;

        internal ScriptShapingContext(OfficeOpenTypeSubstitution owner, string source,
            OfficeTextFeatureSettings settings, string? language, CancellationToken cancellationToken) {
            _owner = owner; Source = source; Settings = settings; Language = language; CancellationToken = cancellationToken;
            Filter = OfficeOpenTypeGlyphFilter.TryCreate(owner._reader.Data)
                ?? throw new InvalidDataException("The font's glyph filtering is unavailable.");
            int firstControl = -1;
            for (int index = 0; index < source.Length; index++) {
                Charge();
                if (firstControl < 0 && IsJoiningControl(source[index])) firstControl = index;
            }
            if (firstControl >= 0) {
                _joiningControls = new int[checked(source.Length + 1)];
                for (int index = firstControl; index < source.Length; index++) {
                    Charge();
                    _joiningControls[index + 1] = _joiningControls[index] + (IsJoiningControl(source[index]) ? 1 : 0);
                }
            }
        }

        internal string Source { get; }
        internal OfficeTextFeatureSettings Settings { get; }
        internal string? Language { get; }
        internal CancellationToken CancellationToken { get; }
        internal OfficeOpenTypeGlyphFilter Filter { get; }
        internal bool HasJoiningControls => _joiningControls != null;
        internal string?[]? Joining => _joining ??= OfficeArabicTextShaper.GetOpenTypeJoiningForms(Source, Charge);

        internal void Charge() {
            CancellationToken.ThrowIfCancellationRequested();
            ChargeOperations(ref Operations, 1);
        }

        internal List<ArabicFeatureLookup>? Lookups(string script) {
            if (!_lookups.TryGetValue(script, out List<ArabicFeatureLookup>? lookups)) {
                lookups = _owner.ReadScriptLookups(script, script == "arab" ? ArabicFeatureStages : MarkScriptFeatureStages, this);
                _lookups.Add(script, lookups);
            }
            return lookups;
        }

        internal bool HasJoiningControl(GlyphToken before, GlyphToken after) {
            Charge();
            if (before.TextIndex < 0 || after.TextIndex > Source.Length || after.TextIndex < before.TextIndex)
                throw new InvalidDataException("GSUB source positions are not in logical order.");
            return _joiningControls != null && _joiningControls[after.TextIndex] != _joiningControls[before.TextIndex];
        }
    }

    internal ScriptShapingContext CreateScriptShapingContext(string source, OfficeTextFeatureSettings settings,
        string? language, CancellationToken cancellationToken) => new(this, source, settings, language, cancellationToken);
}
