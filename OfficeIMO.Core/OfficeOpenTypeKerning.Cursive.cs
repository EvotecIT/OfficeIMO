namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeKerning {
    private sealed partial class LogicalPositioningRun {
        // GPOS type 3 joins the exit anchor of the preceding glyph to the entry anchor of the
        // current one. Along the run direction it rewrites both advances so the anchors meet;
        // across it, the child glyph takes the anchor difference and is chained to its parent,
        // whose own offset it inherits once every lookup has run. Without RightToLeft the later
        // logical glyph is the child; with it, the earlier one is. This follows the reference
        // shaper: https://github.com/harfbuzz/harfbuzz/blob/14.2.1/src/OT/Layout/GPOS/CursivePosFormat1.hh
        private bool ApplyCursive(int subtable, int index, ushort flags, int markSet) {
            Ensure(subtable, 6);
            if (U16(subtable) != 1) throw new InvalidDataException("Unsupported GPOS cursive attachment format.");
            int coverage = Relative(subtable, U16(subtable + 2), 4), records = U16(subtable + 4);
            Ensure(subtable + 6, checked(records * 4));
            int current = Coverage(coverage, _tokens[index].GlyphId);
            if (current < 0) return false;
            if (current >= records) throw new InvalidDataException("GPOS cursive coverage exceeds its records.");
            int entryOffset = U16(subtable + 6 + current * 4 + 0);
            if (entryOffset == 0) return false;
            int previous = NextIncluded(index, -1, flags, markSet, _scripts[index], _tokens.Count - 1);
            if (previous < 0) return false;
            int previousRecord = Coverage(coverage, _tokens[previous].GlyphId);
            if (previousRecord < 0) return false;
            if (previousRecord >= records) throw new InvalidDataException("GPOS cursive coverage exceeds its records.");
            int exitOffset = U16(subtable + 6 + previousRecord * 4 + 2);
            if (exitOffset == 0) return false;
            (int exitX, int exitY) = CursiveAnchor(Relative(subtable, (uint)exitOffset, 6));
            (int entryX, int entryY) = CursiveAnchor(Relative(subtable, (uint)entryOffset, 6));

            int i = previous, j = index;
            if (IsRightToLeft(j)) {
                int d = checked(exitX + _positions[i].XPlacement);
                AddPosition(i, -d, -d, 0, 0);
                SetAdvance(j, checked(entryX + _positions[j].XPlacement));
            } else {
                SetAdvance(i, checked(exitX + _positions[i].XPlacement));
                int d = checked(entryX + _positions[j].XPlacement);
                AddPosition(j, -d, -d, 0, 0);
            }

            int child = i, parent = j, yOffset = checked(entryY - exitY);
            if ((flags & 0x0001) == 0) { child = j; parent = i; yOffset = -yOffset; }
            _cursiveChain ??= new int[_tokens.Count];
            ReverseCursiveChain(child, parent);
            _cursiveChain[child] = parent - child;
            SetYPlacement(child, yOffset);
            if (_cursiveChain[parent] == -_cursiveChain[child]) {
                _cursiveChain[parent] = 0;
                SetYPlacement(parent, 0);
            }
            return true;
        }

        // A child already attached elsewhere moves under its new parent, so its old chain is
        // reversed to hang from it instead of being cut loose.
        private void ReverseCursiveChain(int index, int newParent) {
            int[] chain = _cursiveChain!;
            int guard = 0;
            var path = new List<int>();
            while (chain[index] != 0) {
                if (++guard > chain.Length) throw new InvalidDataException("GPOS cursive attachments form a cycle.");
                _budget.Charge();
                int next = index + chain[index];
                path.Add(index);
                if (next == newParent) break;
                index = next;
            }
            for (int step = path.Count - 1; step >= 0; step--) {
                int from = path[step], to = from + chain[from];
                if (to == newParent) { chain[from] = 0; continue; }
                int offset = chain[from];
                chain[from] = 0;
                SetYPlacement(to, -_positions[from].YPlacement);
                chain[to] = -offset;
            }
        }

        private void PropagateCursiveOffsets() {
            if (_cursiveChain == null) return;
            var resolved = new bool[_tokens.Count];
            for (int index = 0; index < _tokens.Count; index++) ResolveCursiveOffset(index, resolved, 0);
        }

        private void ResolveCursiveOffset(int index, bool[] resolved, int depth) {
            _budget.Charge();
            if (resolved[index]) return;
            if (depth > _tokens.Count) throw new InvalidDataException("GPOS cursive attachments form a cycle.");
            int offset = _cursiveChain![index];
            if (offset != 0) {
                int parent = index + offset;
                ResolveCursiveOffset(parent, resolved, depth + 1);
                AddPosition(index, 0, 0, _positions[parent].YPlacement, 0);
            }
            resolved[index] = true;
        }

        private bool IsRightToLeft(int index) => _rightToLeft != null ? _rightToLeft[index]
            : _scripts[index] is "arab" or "hebr" or "syrc" or "thaa" or "nkoo";

        private void SetAdvance(int index, int advance) {
            int delta = checked(advance - _reader.AdvanceWidth(_tokens[index].GlyphId) - _positions[index].XAdvance);
            AddPosition(index, 0, delta, 0, 0);
        }

        private void SetYPlacement(int index, int y) => AddPosition(index, 0, 0, checked(y - _positions[index].YPlacement), 0);

        // Design-unit anchors; hinting points and device deltas do not apply to unhinted outlines.
        private (int X, int Y) CursiveAnchor(int anchor) {
            int format = U16(anchor);
            if (format == 2) Ensure(anchor, 8);
            else if (format == 3) {
                Ensure(anchor, 10);
                if (U16(anchor + 6) != 0 || U16(anchor + 8) != 0)
                    throw new InvalidDataException("GPOS device or variable cursive anchors require a positioned-outline provider.");
            } else if (format != 1) throw new InvalidDataException("Unsupported GPOS anchor format.");
            return (unchecked((short)U16(anchor + 2)), unchecked((short)U16(anchor + 4)));
        }
    }
}
