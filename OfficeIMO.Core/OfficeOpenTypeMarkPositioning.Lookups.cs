using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeMarkPositioning {
    private void ApplyLookup(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens, string[] scripts,
        OfficeOpenTypeMarkAttachment[] result, bool[] interpretedMarks, int lookupIndex, HashSet<string> activeScripts) {
        int count = Count(_lookupList, 2);
        if (lookupIndex < 0 || lookupIndex >= count) throw new InvalidDataException("GPOS lookup index is invalid.");
        int lookup = Relative(_lookupList, U16(_lookupList + 2 + lookupIndex * 2), 6);
        int type = U16(lookup), subtableCount = U16(lookup + 4);
        ushort flags = U16(lookup + 2);
        if (subtableCount > 256) throw new InvalidDataException("GPOS subtable count exceeds the managed limit.");
        Ensure(lookup + 6, checked(subtableCount * 2));
        int markSet = -1;
        if ((flags & 0x0010) != 0) markSet = U16(lookup + 6 + subtableCount * 2);
        if (!_filter.SupportsFlags(flags, markSet)) throw new InvalidDataException("Unsupported GPOS mark filtering.");
        var subtables = new List<(int Type, int Offset)>();
        for (int index = 0; index < subtableCount; index++) {
            Charge(); int subtable = Relative(lookup, U16(lookup + 6 + index * 2), 2), effectiveType = type;
            if (type == 9) {
                Ensure(subtable, 8);
                if (U16(subtable) != 1) throw new InvalidDataException("Unsupported GPOS extension format.");
                effectiveType = U16(subtable + 2);
                subtable = Relative(subtable, U32(subtable + 4), 12);
            }
            if (effectiveType is < 4 or > 6) throw new InvalidDataException("The selected GPOS feature requires an unsupported mark lookup.");
            Ensure(subtable, 12);
            if (U16(subtable) != 1) throw new InvalidDataException("Unsupported GPOS mark attachment format.");
            subtables.Add((effectiveType, subtable));
        }
        for (int index = 0; index < tokens.Count; index++) {
            Charge();
            if (!result[index].IsMark || !activeScripts.Contains(scripts[index])
                || _filter.IsIgnored(tokens[index], flags, markSet, Charge)) continue;
            foreach ((int effectiveType, int subtable) in subtables) {
                if (TryAttach(tokens, result, interpretedMarks, index, effectiveType, subtable, flags, markSet)) break;
            }
        }
    }

    private bool TryAttach(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        OfficeOpenTypeMarkAttachment[] result, bool[] interpretedMarks, int index, int type, int subtable, ushort flags, int markSet) {
        int markCoverage = Relative(subtable, U16(subtable + 2), 4);
        int markIndex = Coverage(markCoverage, tokens[index].GlyphId);
        if (markIndex < 0) return false;
        int classCount = U16(subtable + 6);
        if (classCount < 1 || classCount > MaximumRecords) throw new InvalidDataException("GPOS mark class count exceeds the managed limit.");
        int markArray = Relative(subtable, U16(subtable + 8), 2), marks = U16(markArray);
        Ensure(markArray + 2, checked(marks * 4));
        if (markIndex >= marks) throw new InvalidDataException("GPOS mark coverage exceeds its mark array.");
        int record = markArray + 2 + markIndex * 4, markClass = U16(record);
        if (markClass >= classCount) throw new InvalidDataException("GPOS mark class is invalid.");
        int parentCoverage = Relative(subtable, U16(subtable + 4), 4);
        _ = Coverage(parentCoverage, -1); // Validate the coverage format even when no parent matches.
        int parentArray = Relative(subtable, U16(subtable + 10), 2), parents = U16(parentArray);
        Ensure(parentArray + 2, checked(parents * (type == 5 ? 2 : classCount * 2)));
        int markAnchor = Optional(markArray, U16(record + 2), 6);
        // Validate the selected mark anchor before considering a missing parent;
        // unsupported device/variation anchors cannot masquerade as a no-op.
        (int markX, int markY) = markAnchor == 0 ? (0, 0) : Anchor(markAnchor);
        interpretedMarks[index] = true;
        if (markAnchor == 0) return false;
        int parent = FindParent(tokens, result, index, type, flags, markSet);
        if (parent < 0) return false;
        int parentIndex = Coverage(parentCoverage, tokens[parent].GlyphId);
        if (parentIndex < 0) return false;
        if (parentIndex >= parents) throw new InvalidDataException("GPOS parent coverage exceeds its attachment array.");
        int anchor;
        if (type == 5) {
            Ensure(parentArray + 2, checked(parents * 2));
            int ligature = Relative(parentArray, U16(parentArray + 2 + parentIndex * 2), 2);
            int components = U16(ligature);
            if (components < 1 || components > 256) throw new InvalidDataException("GPOS ligature component count exceeds the managed limit.");
            Ensure(ligature + 2, checked(components * classCount * 2));
            int component = LigatureComponent(tokens[parent], tokens[index], components);
            anchor = Optional(ligature, U16(ligature + 2 + (component * classCount + markClass) * 2), 6);
        } else {
            Ensure(parentArray + 2, checked(parents * classCount * 2));
            anchor = Optional(parentArray, U16(parentArray + 2 + (parentIndex * classCount + markClass) * 2), 6);
        }
        if (anchor == 0) return false;
        (int baseX, int baseY) = Anchor(anchor);
        result[index] = new OfficeOpenTypeMarkAttachment(parent, checked(baseX - markX), checked(baseY - markY), isMark: true);
        return true;
    }

    private int FindParent(IReadOnlyList<OfficeOpenTypeSubstitution.GlyphToken> tokens,
        OfficeOpenTypeMarkAttachment[] result, int index, int type, ushort flags, int markSet) {
        for (int previous = index - 1; previous >= 0; previous--) {
            Charge();
            // A retained default-ignorable preserves logical source without
            // introducing a glyph origin to which visible marks may attach.
            if (OfficeManagedTextShaper.IsNonPaintingControl(tokens[previous].Scalar)) continue;
            if (type == 6) {
                // Parent marks retain attachment-class/filter-set selection, but
                // ordinary ignore flags must not let a search cross a base.
                ushort parentFlags = (ushort)(flags & ~0x000E);
                if (_filter.IsIgnored(tokens[previous], parentFlags, markSet, Charge)) continue;
                return result[previous].IsMark && SharesLigatureComponent(tokens[index], tokens[previous]) ? previous : -1;
            }
            // A mark-to-base/ligature lookup searches through intervening marks.
            // Its input flags filter the mark being positioned, not the parent.
            if (result[previous].IsMark) continue;
            // Coverage, rather than a possibly incomplete GDEF classification, owns
            // whether this preceding nonmark has a base/ligature attachment record.
            return previous;
        }
        return -1;
    }

    private static int LigatureComponent(OfficeOpenTypeSubstitution.GlyphToken ligature,
        OfficeOpenTypeSubstitution.GlyphToken mark, int componentCount) {
        // Components in GPOS can include a precomposed source glyph's own parts:
        // Noto Arabic's lam + alef-with-hamza has two GSUB inputs but three
        // anchor rows. Explicit GSUB association is independent of Unicode text
        // ownership: normalization can assign a whole grapheme to one token.
        return ligature.LigatureIdentity != null && ReferenceEquals(ligature.LigatureIdentity, mark.LigatureIdentity)
            && mark.LigatureComponent > 0 ? Math.Min(mark.LigatureComponent, componentCount) - 1 : componentCount - 1;
    }

    private static bool SharesLigatureComponent(OfficeOpenTypeSubstitution.GlyphToken first,
        OfficeOpenTypeSubstitution.GlyphToken second) {
        if (ReferenceEquals(first.LigatureIdentity, second.LigatureIdentity))
            return first.LigatureIdentity == null || first.LigatureComponent == second.LigatureComponent;
        // A mark can itself be a GSUB ligature. It remains a valid attachment
        // parent despite having its own identity and no base-component index.
        return first.LigatureIdentity != null && first.LigatureComponent == 0
            || second.LigatureIdentity != null && second.LigatureComponent == 0;
    }

    private (int X, int Y) Anchor(int anchor) {
        int format = U16(anchor);
        Ensure(anchor, 6);
        if (format == 2) Ensure(anchor, 8); // Unhinted outlines use the supplied design coordinates.
        else if (format == 3) {
            Ensure(anchor, 10);
            if (U16(anchor + 6) != 0 || U16(anchor + 8) != 0)
                throw new InvalidDataException("GPOS device or variable mark anchors require a positioned-outline provider.");
        } else if (format != 1) throw new InvalidDataException("Unsupported GPOS anchor format.");
        return (S16(anchor + 2), S16(anchor + 4));
    }

    private int Coverage(int table, int glyph) {
        Ensure(table, 4);
        int format = U16(table), count = U16(table + 2), stride = format == 1 ? 2 : format == 2 ? 6 : 0;
        if (stride == 0) throw new InvalidDataException("Unsupported GPOS coverage format.");
        Ensure(table + 4, checked(count * stride));
        int low = 0, high = count - 1;
        while (low <= high) {
            Charge(); int middle = low + ((high - low) / 2), record = table + 4 + middle * stride;
            int first = U16(record), last = format == 1 ? first : U16(record + 2);
            if (glyph < first) high = middle - 1;
            else if (glyph > last) low = middle + 1;
            else return format == 1 ? middle : checked(U16(record + 4) + glyph - first);
        }
        return -1;
    }
}
