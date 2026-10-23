using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeKerning {
    private sealed partial class LogicalPositioningRun {
        private bool ApplyLookup(ushort lookupIndex, int index, int inputEnd, int depth, out int next) {
            _budget.Charge(); next = index + 1;
            if (depth >= MaximumNesting) throw new InvalidDataException("GPOS contextual lookup nesting limit exceeded.");
            int lookup = Relative(_lookupList, U16(_lookupList + 2 + lookupIndex * 2), 6);
            ushort type = U16(lookup), flags = U16(lookup + 2);
            int count = U16(lookup + 4);
            if (count > MaximumRecords) throw new InvalidDataException("GPOS subtable limit exceeded.");
            Ensure(lookup + 6, checked(count * 2));
            int markSet = (flags & 0x10) != 0 ? U16(lookup + 6 + count * 2) : 0;
            if (!_filter.SupportsFlags(flags, markSet)) throw new InvalidDataException("Unsupported GPOS glyph filtering.");
            if (_filter.IsIgnored(_tokens[index], flags, markSet, _budget.Charge)) return false;
            for (int subtableIndex = 0; subtableIndex < count; subtableIndex++) {
                _budget.Charge();
                int subtable = Relative(lookup, U16(lookup + 6 + subtableIndex * 2), 2);
                ushort effectiveType = type;
                if (type == 9) {
                    Ensure(subtable, 8);
                    if (!_owner._includeExtendedGpos || U16(subtable) != 1)
                        throw new InvalidDataException("Unsupported GPOS extension.");
                    effectiveType = U16(subtable + 2);
                    subtable = Relative(subtable, U32(subtable + 4), 2);
                }
                bool matched;
                switch (effectiveType) {
                    case 1:
                        matched = ApplySingle(subtable, index);
                        break;
                    case 2:
                        matched = ApplyPair(subtable, index, inputEnd, flags, markSet, out next);
                        break;
                    case 8:
                        matched = ApplyChainedContext(subtable, index, inputEnd, flags, markSet, depth, out next);
                        break;
                    default:
                        throw new InvalidDataException("Unsupported active GPOS kern lookup.");
                }
                if (matched) return true;
            }
            return false;
        }

        private bool ApplySingle(int subtable, int index) {
            Ensure(subtable, 6);
            int format = U16(subtable);
            if (format != 1 && format != 2) throw new InvalidDataException("Unsupported GPOS single-adjustment format.");
            int coverage = Relative(subtable, U16(subtable + 2), 4);
            int coverageIndex = Coverage(coverage, _tokens[index].GlyphId);
            if (coverageIndex < 0) return false;
            ushort valueFormat = U16(subtable + 4);
            int offset = subtable + 6;
            if (format == 2) {
                int count = U16(subtable + 6);
                if (coverageIndex >= count) throw new InvalidDataException("GPOS single adjustment coverage exceeds its value array.");
                Ensure(subtable + 8, checked(count * ValueRecordSize(valueFormat)));
                offset = checked(subtable + 8 + coverageIndex * ValueRecordSize(valueFormat));
            }
            OfficeOpenTypePairValue value = _owner.ReadValueRecord(subtable, offset, valueFormat, _budget);
            _positions[index] = _positions[index].Add(value.XPlacement, value.XAdvance, value.YPlacement, value.YAdvance);
            return true;
        }

        private bool ApplyPair(int subtable, int index, int inputEnd, ushort flags, int markSet, out int next) {
            next = index + 1;
            Ensure(subtable, 10);
            int format = U16(subtable);
            if (format != 1 && format != 2 || format == 2 && !_owner._includeExtendedGpos)
                throw new InvalidDataException("Unsupported GPOS pair-adjustment format.");
            int second = NextIncluded(index, 1, flags, markSet, _scripts[index], inputEnd);
            if (second < 0) return false;
            int coverage = Relative(subtable, U16(subtable + 2), 4);
            if (Coverage(coverage, _tokens[index].GlyphId) < 0) return false;
            // The canonical pair reader owns both formats and all value/variation fields.
            if (!_owner.TryGposPairAdjustmentFromSubtable(subtable, (ushort)_tokens[index].GlyphId,
                (ushort)_tokens[second].GlyphId, _budget, out OfficeOpenTypePairPositioning adjustment)) return false;
            _positions[index] = _positions[index].Add(adjustment.FirstGlyphXPlacement,
                adjustment.FirstGlyphXAdvance, adjustment.FirstGlyphYPlacement, adjustment.FirstGlyphYAdvance);
            _positions[second] = _positions[second].Add(adjustment.SecondGlyphXPlacement,
                adjustment.SecondGlyphXAdvance, adjustment.SecondGlyphYPlacement, adjustment.SecondGlyphYAdvance);
            next = adjustment.SecondValueRecordPresent ? second + 1 : second;
            return true;
        }

        private bool ApplyChainedContext(int subtable, int index, int inputEnd, ushort flags,
            int markSet, int depth, out int next) {
            next = index + 1;
            if (U16(subtable) != 3) throw new InvalidDataException("Unsupported GPOS chained-context format.");
            int cursor = subtable + 2;
            int backCount = ContextCount(ref cursor), backCoverages = cursor;
            cursor = checked(cursor + backCount * 2);
            int inputCount = ContextCount(ref cursor), inputCoverages = cursor;
            if (inputCount == 0) throw new InvalidDataException("An empty GPOS input context is invalid.");
            cursor = checked(cursor + inputCount * 2);
            int lookaheadCount = ContextCount(ref cursor), lookaheadCoverages = cursor;
            cursor = checked(cursor + lookaheadCount * 2);
            int recordCount = U16(cursor), records = cursor + 2;
            if (recordCount > MaximumRecords) throw new InvalidDataException("GPOS contextual record limit exceeded.");
            Ensure(records, checked(recordCount * 4));
            string script = _scripts[index];
            var inputs = new int[inputCount];
            int candidate = index;
            for (int offset = 0; offset < inputCount; offset++) {
                _budget.Charge();
                if (offset > 0) candidate = NextIncluded(candidate, 1, flags, markSet, script, inputEnd);
                if (candidate < 0 || Coverage(Relative(subtable, U16(inputCoverages + offset * 2), 4), _tokens[candidate].GlyphId) < 0)
                    return false;
                inputs[offset] = candidate;
            }
            candidate = index;
            for (int offset = 0; offset < backCount; offset++) {
                candidate = NextIncluded(candidate, -1, flags, markSet, script, _tokens.Count - 1);
                if (candidate < 0 || Coverage(Relative(subtable, U16(backCoverages + offset * 2), 4), _tokens[candidate].GlyphId) < 0)
                    return false;
            }
            candidate = inputs[inputCount - 1];
            for (int offset = 0; offset < lookaheadCount; offset++) {
                candidate = NextIncluded(candidate, 1, flags, markSet, script, _tokens.Count - 1);
                if (candidate < 0 || Coverage(Relative(subtable, U16(lookaheadCoverages + offset * 2), 4), _tokens[candidate].GlyphId) < 0)
                    return false;
            }
            for (int offset = 0; offset < recordCount; offset++) {
                _budget.Charge();
                int record = records + offset * 4, sequenceIndex = U16(record), nestedLookup = U16(record + 2);
                if (sequenceIndex >= inputCount || nestedLookup >= U16(_lookupList))
                    throw new InvalidDataException("Invalid GPOS contextual positioning record.");
                ApplyLookup((ushort)nestedLookup, inputs[sequenceIndex], inputs[inputCount - 1], depth + 1, out _);
            }
            next = inputs[inputCount - 1] + 1;
            return true;
        }

        private int NextIncluded(int index, int step, ushort flags, int markSet, string script, int inputEnd) {
            for (int next = index + step; next >= 0 && next <= inputEnd; next += step) {
                _budget.Charge();
                if (_scripts[next] != script) return -1;
                if (!_filter.IsIgnored(_tokens[next], flags, markSet, _budget.Charge)) return next;
            }
            return -1;
        }

        private int ContextCount(ref int cursor) {
            int count = U16(cursor); cursor += 2;
            if (count > MaximumContextGlyphs) throw new InvalidDataException("GPOS context length exceeds the managed limit.");
            Ensure(cursor, checked(count * 2)); return count;
        }

        private int Coverage(int table, int glyph) {
            int format = U16(table), count = U16(table + 2);
            int stride = format == 1 ? 2 : format == 2 ? 6 : 0;
            if (stride == 0) throw new InvalidDataException("Unsupported GPOS coverage format.");
            Ensure(table + 4, checked(count * stride));
            return _owner.CoverageIndex(table, (ushort)glyph, _budget);
        }
    }
}
