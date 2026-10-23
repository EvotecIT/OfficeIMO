using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    private object? ClassifyLigatureInputs(List<GlyphToken> glyphs, int first, int count,
        List<int> componentStarts, ref int operations) {
        OfficeOpenTypeGlyphFilter? filter = OfficeOpenTypeGlyphFilter.TryCreate(_reader.Data);
        if (filter == null) throw new InvalidDataException("A GSUB ligature has no usable glyph classification.");
        var inputs = new List<LigatureAssociationInput>();
        int inspections = operations;
        Action charge = () => ChargeOperations(ref inspections, 1);
        for (int component = 0; component < count; component++) {
            GlyphToken token = glyphs[first + component];
            if (filter.GetGlyphClass(token.GlyphId, token.Scalar, token.LigatureComponentTextIndexes?.Count > 1, charge) == 3) continue;
            int offset = componentStarts.Count;
            if (token.LigatureComponentTextIndexes?.Count > 0) componentStarts.AddRange(token.LigatureComponentTextIndexes);
            else componentStarts.Add(token.TextIndex);
            inputs.Add(new LigatureAssociationInput(token.LigatureIdentity, token.LookupOrder,
                offset, componentStarts.Count - offset));
        }
        operations = inspections;
        return inputs.Count > 1 ? new LigatureAssociation(glyphs[first].LookupIdentity, inputs.ToArray()) : glyphs[first].LigatureIdentity;
    }

    /// <summary>One non-mark input's contribution to a font-designed ligature.</summary>
    private readonly struct LigatureAssociationInput {
        internal LigatureAssociationInput(object? identity, int order, int componentOffset, int componentCount) {
            Identity = identity; Order = order; ComponentOffset = componentOffset; ComponentCount = componentCount;
        }
        internal object? Identity { get; }
        internal int Order { get; }
        internal int ComponentOffset { get; }
        internal int ComponentCount { get; }
    }

    /// <summary>Identity and bounded remapping of marks retained within nested GSUB inputs.</summary>
    /// <remarks>Lookup ordinals distinguish normalized continuations sharing one original source index.</remarks>
    private sealed class LigatureAssociation {
        private readonly object? _lookupIdentity;
        private readonly LigatureAssociationInput[] _inputs;

        internal LigatureAssociation(object? lookupIdentity, LigatureAssociationInput[] inputs) {
            _lookupIdentity = lookupIdentity; _inputs = inputs;
        }

        internal bool TryAssociate(GlyphToken mark, ref int operations, out int component) {
            if (mark.LigatureIdentity != null && mark.LigatureComponent > 0
                && TryRemap(mark.LigatureIdentity, mark.LigatureComponent, 0, ref operations, out component)) return true;
            if (!ReferenceEquals(mark.LookupIdentity, _lookupIdentity)) { component = 0; return false; }
            return TryAssociateOrder(mark.LookupOrder, 0, ref operations, out component);
        }

        private bool TryRemap(object identity, int component, int depth, ref int operations, out int mapped) {
            ChargeOperations(ref operations, 1);
            if (depth >= MaximumLookupRecursion) throw new InvalidDataException("GSUB ligature association nesting exceeds the managed limit.");
            if (ReferenceEquals(identity, this)) { mapped = component; return true; }
            foreach (LigatureAssociationInput input in _inputs) {
                ChargeOperations(ref operations, 1);
                if (ReferenceEquals(identity, input.Identity)) {
                    mapped = checked(input.ComponentOffset + Math.Min(component, input.ComponentCount)); return true;
                }
                if (input.Identity is LigatureAssociation previous
                    && previous.TryRemap(identity, component, depth + 1, ref operations, out int inherited)) {
                    mapped = checked(input.ComponentOffset + Math.Min(inherited, input.ComponentCount)); return true;
                }
            }
            mapped = 0; return false;
        }

        private bool TryAssociateOrder(int order, int depth, ref int operations, out int component) {
            ChargeOperations(ref operations, 1);
            if (depth >= MaximumLookupRecursion) throw new InvalidDataException("GSUB ligature association nesting exceeds the managed limit.");
            if (_inputs.Length < 2 || order <= _inputs[0].Order || order >= _inputs[_inputs.Length - 1].Order) {
                component = 0; return false;
            }
            for (int index = _inputs.Length - 2; index >= 0; index--) {
                ChargeOperations(ref operations, 1);
                LigatureAssociationInput input = _inputs[index];
                if (order <= input.Order) continue;
                int local = input.ComponentCount;
                if (input.Identity is LigatureAssociation previous && ReferenceEquals(previous._lookupIdentity, _lookupIdentity)
                    && previous.TryAssociateOrder(order, depth + 1, ref operations, out int nested)) local = nested;
                component = checked(input.ComponentOffset + local); return true;
            }
            component = 0; return false;
        }
    }
}
