using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Unicode Line Breaking Algorithm (UAX #14, rules LB1–LB31) over Unicode scalars.</summary>
internal static class OfficeUnicodeLineBreaker {
    /// <summary>Break action at the boundary before a scalar.</summary>
    internal enum BreakAction : byte {
        Prohibited,
        Allowed,
        Mandatory
    }

    /// <summary>
    /// Returns the action at each boundary: element <c>i</c> describes the boundary before scalar
    /// <c>i</c>. Element 0 (start of text) is always prohibited; the end of text is not included.
    /// </summary>
    internal static BreakAction[] Analyze(int[] scalars, System.Threading.CancellationToken cancellationToken = default) {
        int count = scalars.Length;
        var actions = new BreakAction[count];
        if (count == 0) return actions;

        var raw = new OfficeLineBreakClass[count];
        var classes = new OfficeLineBreakClass[count];
        var attached = new bool[count];
        for (int index = 0; index < count; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            raw[index] = ResolveLb1(scalars[index]);
        }
        for (int index = 0; index < count; index++) {
            OfficeLineBreakClass type = raw[index];
            if (type is OfficeLineBreakClass.CM or OfficeLineBreakClass.ZWJ) {
                // LB9: X (CM|ZWJ)* is treated as X; LB10: remaining CM/ZWJ are AL.
                if (index > 0 && !IsLb9Excluded(classes[index - 1])) {
                    classes[index] = classes[index - 1];
                    attached[index] = true;
                    continue;
                }
                classes[index] = OfficeLineBreakClass.AL;
                continue;
            }
            classes[index] = type;
        }

        var context = new Context(scalars, raw, classes, attached);
        for (int index = 1; index < count; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            actions[index] = context.Resolve(index);
        }
        return actions;
    }

    private static OfficeLineBreakClass ResolveLb1(int scalar) {
        OfficeLineBreakClass type = OfficeUnicodeTextData.GetLineBreakClass(scalar);
        switch (type) {
            case OfficeLineBreakClass.AI:
            case OfficeLineBreakClass.SG:
            case OfficeLineBreakClass.XX:
                return OfficeLineBreakClass.AL;
            case OfficeLineBreakClass.CJ:
                return OfficeLineBreakClass.NS;
            case OfficeLineBreakClass.SA:
                UnicodeCategory category = GetCategory(scalar);
                return category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark
                    ? OfficeLineBreakClass.CM : OfficeLineBreakClass.AL;
            default:
                return type;
        }
    }

    private static bool IsLb9Excluded(OfficeLineBreakClass type) =>
        type is OfficeLineBreakClass.BK or OfficeLineBreakClass.CR or OfficeLineBreakClass.LF or
            OfficeLineBreakClass.NL or OfficeLineBreakClass.SP or OfficeLineBreakClass.ZW;

    internal static UnicodeCategory GetCategory(int scalar) =>
        scalar > 0xFFFF
            ? CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0)
            : CharUnicodeInfo.GetUnicodeCategory((char)scalar);

    private sealed class Context {
        private readonly int[] _scalars;
        private readonly OfficeLineBreakClass[] _raw;
        private readonly OfficeLineBreakClass[] _classes;
        private readonly bool[] _attached;
        private readonly int[] _regionalCounts;

        internal Context(int[] scalars, OfficeLineBreakClass[] raw, OfficeLineBreakClass[] classes, bool[] attached) {
            _scalars = scalars;
            _raw = raw;
            _classes = classes;
            _attached = attached;
            _regionalCounts = new int[classes.Length];
            int regional = 0;
            for (int index = 0; index < classes.Length; index++) {
                if (!attached[index]) regional = classes[index] == OfficeLineBreakClass.RI ? regional + 1 : 0;
                _regionalCounts[index] = regional;
            }
        }

        // Previous position that is not part of an earlier position's combining sequence.
        private int Previous(int index) {
            index--;
            while (index >= 0 && _attached[index]) index--;
            return index;
        }

        private int Next(int index) {
            index++;
            while (index < _classes.Length && _attached[index]) index++;
            return index;
        }

        private OfficeLineBreakClass ClassAt(int index) => _classes[index];
        private bool IsEastAsian(int index) => OfficeUnicodeTextData.IsEastAsianWide(_scalars[index]);
        private bool IsQuote(int index, UnicodeCategory category) =>
            _classes[index] == OfficeLineBreakClass.QU && GetCategory(_scalars[index]) == category;
        private bool IsAksaraBase(int index) =>
            _classes[index] is OfficeLineBreakClass.AK or OfficeLineBreakClass.AS || _scalars[index] == 0x25CC;

        internal BreakAction Resolve(int index) {
            int beforeIndex = index - 1;
            OfficeLineBreakClass before = _classes[beforeIndex];
            OfficeLineBreakClass rawBefore = _raw[beforeIndex];
            OfficeLineBreakClass after = _classes[index];
            OfficeLineBreakClass rawAfter = _raw[index];

            // LB4, LB5
            if (rawBefore == OfficeLineBreakClass.BK) return BreakAction.Mandatory;
            if (rawBefore == OfficeLineBreakClass.CR && rawAfter == OfficeLineBreakClass.LF) return BreakAction.Prohibited;
            if (rawBefore is OfficeLineBreakClass.CR or OfficeLineBreakClass.LF or OfficeLineBreakClass.NL) return BreakAction.Mandatory;
            // LB6
            if (rawAfter is OfficeLineBreakClass.BK or OfficeLineBreakClass.CR or OfficeLineBreakClass.LF or OfficeLineBreakClass.NL)
                return BreakAction.Prohibited;
            // LB7
            if (rawAfter is OfficeLineBreakClass.SP or OfficeLineBreakClass.ZW) return BreakAction.Prohibited;
            // LB8: ZW SP* ÷
            int lastNonSpace = beforeIndex;
            while (lastNonSpace >= 0 && _raw[lastNonSpace] == OfficeLineBreakClass.SP) lastNonSpace--;
            if (lastNonSpace >= 0 && _raw[lastNonSpace] == OfficeLineBreakClass.ZW) return BreakAction.Allowed;
            // LB8a
            if (rawBefore == OfficeLineBreakClass.ZWJ) return BreakAction.Prohibited;
            // LB9
            if (_attached[index]) return BreakAction.Prohibited;

            // Rules below see combining sequences as their base character.
            int b = _attached[beforeIndex] ? Previous(index) : beforeIndex;
            before = ClassAt(b);
            int spaceRunStart = b;
            while (spaceRunStart >= 0 && _classes[spaceRunStart] == OfficeLineBreakClass.SP) spaceRunStart = Previous(spaceRunStart);
            OfficeLineBreakClass beforeSpaces = spaceRunStart >= 0 ? _classes[spaceRunStart] : OfficeLineBreakClass.XX;
            int next = Next(index);
            bool atEnd = next >= _classes.Length;
            OfficeLineBreakClass afterNext = atEnd ? OfficeLineBreakClass.XX : _classes[next];

            // LB11
            if (after == OfficeLineBreakClass.WJ || before == OfficeLineBreakClass.WJ) return BreakAction.Prohibited;
            // LB12, LB12a
            if (before == OfficeLineBreakClass.GL) return BreakAction.Prohibited;
            if (after == OfficeLineBreakClass.GL && before is not (OfficeLineBreakClass.SP or OfficeLineBreakClass.BA or OfficeLineBreakClass.HY))
                return BreakAction.Prohibited;
            // LB13
            if (after is OfficeLineBreakClass.CL or OfficeLineBreakClass.CP or OfficeLineBreakClass.EX or OfficeLineBreakClass.SY)
                return BreakAction.Prohibited;
            // LB14: OP SP* ×
            if (beforeSpaces == OfficeLineBreakClass.OP) return BreakAction.Prohibited;
            // LB15a: (sot|BK|CR|LF|NL|OP|QU|GL|SP|ZW) [Pi&QU] SP* ×
            if (spaceRunStart >= 0 && IsQuote(spaceRunStart, UnicodeCategory.InitialQuotePunctuation)) {
                int prior = Previous(spaceRunStart);
                if (prior < 0 || _classes[prior] is OfficeLineBreakClass.BK or OfficeLineBreakClass.CR or OfficeLineBreakClass.LF or
                    OfficeLineBreakClass.NL or OfficeLineBreakClass.OP or OfficeLineBreakClass.QU or OfficeLineBreakClass.GL or
                    OfficeLineBreakClass.SP or OfficeLineBreakClass.ZW) return BreakAction.Prohibited;
            }
            // LB15b: × [Pf&QU] (SP|GL|WJ|CL|QU|CP|EX|IS|SY|BK|CR|LF|NL|ZW|eot)
            if (IsQuote(index, UnicodeCategory.FinalQuotePunctuation) &&
                (atEnd || afterNext is OfficeLineBreakClass.SP or OfficeLineBreakClass.GL or OfficeLineBreakClass.WJ or
                    OfficeLineBreakClass.CL or OfficeLineBreakClass.QU or OfficeLineBreakClass.CP or OfficeLineBreakClass.EX or
                    OfficeLineBreakClass.IS or OfficeLineBreakClass.SY or OfficeLineBreakClass.BK or OfficeLineBreakClass.CR or
                    OfficeLineBreakClass.LF or OfficeLineBreakClass.NL or OfficeLineBreakClass.ZW)) return BreakAction.Prohibited;
            // LB15c, LB15d
            if (before == OfficeLineBreakClass.SP && after == OfficeLineBreakClass.IS && afterNext == OfficeLineBreakClass.NU && !atEnd)
                return BreakAction.Allowed;
            if (after == OfficeLineBreakClass.IS) return BreakAction.Prohibited;
            // LB16, LB17
            if (after == OfficeLineBreakClass.NS && beforeSpaces is OfficeLineBreakClass.CL or OfficeLineBreakClass.CP) return BreakAction.Prohibited;
            if (after == OfficeLineBreakClass.B2 && beforeSpaces == OfficeLineBreakClass.B2) return BreakAction.Prohibited;
            // LB18
            if (before == OfficeLineBreakClass.SP) return BreakAction.Allowed;
            // LB19
            if (after == OfficeLineBreakClass.QU && GetCategory(_scalars[index]) != UnicodeCategory.InitialQuotePunctuation) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.QU && GetCategory(_scalars[b]) != UnicodeCategory.FinalQuotePunctuation) return BreakAction.Prohibited;
            // LB19a
            if (after == OfficeLineBreakClass.QU) {
                if (!IsEastAsian(b)) return BreakAction.Prohibited;
                if (atEnd || !IsEastAsian(next)) return BreakAction.Prohibited;
            }
            if (before == OfficeLineBreakClass.QU) {
                if (!IsEastAsian(index)) return BreakAction.Prohibited;
                int prior = Previous(b);
                if (prior < 0 || !IsEastAsian(prior)) return BreakAction.Prohibited;
            }
            // LB20
            if (after == OfficeLineBreakClass.CB || before == OfficeLineBreakClass.CB) return BreakAction.Allowed;
            // LB20a
            if ((before == OfficeLineBreakClass.HY || _scalars[b] == 0x2010) && after == OfficeLineBreakClass.AL) {
                int prior = Previous(b);
                if (prior < 0 || _classes[prior] is OfficeLineBreakClass.BK or OfficeLineBreakClass.CR or OfficeLineBreakClass.LF or
                    OfficeLineBreakClass.NL or OfficeLineBreakClass.SP or OfficeLineBreakClass.ZW or OfficeLineBreakClass.CB or
                    OfficeLineBreakClass.GL) return BreakAction.Prohibited;
            }
            // LB21
            if (after is OfficeLineBreakClass.BA or OfficeLineBreakClass.HY or OfficeLineBreakClass.NS || before == OfficeLineBreakClass.BB)
                return BreakAction.Prohibited;
            // LB21a: HL (HY | [BA - EastAsian]) × [^HL]
            if (after != OfficeLineBreakClass.HL && (before == OfficeLineBreakClass.HY || before == OfficeLineBreakClass.BA && !IsEastAsian(b))) {
                int prior = Previous(b);
                if (prior >= 0 && _classes[prior] == OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            }
            // LB21b, LB22
            if (before == OfficeLineBreakClass.SY && after == OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            if (after == OfficeLineBreakClass.IN) return BreakAction.Prohibited;
            // LB23, LB23a, LB24
            if (before is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL && after == OfficeLineBreakClass.NU) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.NU && after is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.PR && after is OfficeLineBreakClass.ID or OfficeLineBreakClass.EB or OfficeLineBreakClass.EM) return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.ID or OfficeLineBreakClass.EB or OfficeLineBreakClass.EM && after == OfficeLineBreakClass.PO) return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.PR or OfficeLineBreakClass.PO && after is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL && after is OfficeLineBreakClass.PR or OfficeLineBreakClass.PO) return BreakAction.Prohibited;
            // LB25
            if (ResolveNumber(b, before, index, after, next, atEnd, afterNext)) return BreakAction.Prohibited;
            // LB26, LB27
            if (before == OfficeLineBreakClass.JL && after is OfficeLineBreakClass.JL or OfficeLineBreakClass.JV or OfficeLineBreakClass.H2 or OfficeLineBreakClass.H3)
                return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.JV or OfficeLineBreakClass.H2 && after is OfficeLineBreakClass.JV or OfficeLineBreakClass.JT) return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.JT or OfficeLineBreakClass.H3 && after == OfficeLineBreakClass.JT) return BreakAction.Prohibited;
            if (before is OfficeLineBreakClass.JL or OfficeLineBreakClass.JV or OfficeLineBreakClass.JT or OfficeLineBreakClass.H2 or OfficeLineBreakClass.H3 &&
                after == OfficeLineBreakClass.PO) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.PR &&
                after is OfficeLineBreakClass.JL or OfficeLineBreakClass.JV or OfficeLineBreakClass.JT or OfficeLineBreakClass.H2 or OfficeLineBreakClass.H3)
                return BreakAction.Prohibited;
            // LB28
            if (before is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL && after is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            // LB28a
            if (before == OfficeLineBreakClass.AP && IsAksaraBase(index)) return BreakAction.Prohibited;
            if (IsAksaraBase(b) && after is OfficeLineBreakClass.VF or OfficeLineBreakClass.VI) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.VI && (after == OfficeLineBreakClass.AK || _scalars[index] == 0x25CC)) {
                int prior = Previous(b);
                if (prior >= 0 && IsAksaraBase(prior)) return BreakAction.Prohibited;
            }
            if (IsAksaraBase(b) && IsAksaraBase(index) && !atEnd && afterNext == OfficeLineBreakClass.VF) return BreakAction.Prohibited;
            // LB29
            if (before == OfficeLineBreakClass.IS && after is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL) return BreakAction.Prohibited;
            // LB30
            if (before is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL or OfficeLineBreakClass.NU &&
                after == OfficeLineBreakClass.OP && !IsEastAsian(index)) return BreakAction.Prohibited;
            if (before == OfficeLineBreakClass.CP && !IsEastAsian(b) &&
                after is OfficeLineBreakClass.AL or OfficeLineBreakClass.HL or OfficeLineBreakClass.NU) return BreakAction.Prohibited;
            // LB30a
            if (before == OfficeLineBreakClass.RI && after == OfficeLineBreakClass.RI) {
                if ((_regionalCounts[b] & 1) == 1) return BreakAction.Prohibited;
            }
            // LB30b
            if (after == OfficeLineBreakClass.EM) {
                if (before == OfficeLineBreakClass.EB) return BreakAction.Prohibited;
                if (OfficeUnicodeTextData.IsExtendedPictographic(_scalars[b]) && GetCategory(_scalars[b]) == UnicodeCategory.OtherNotAssigned)
                    return BreakAction.Prohibited;
            }
            // LB31
            return BreakAction.Allowed;
        }

        // LB25 numeric expressions (Unicode 15.1+ regular-expression form).
        private bool ResolveNumber(int b, OfficeLineBreakClass before, int index, OfficeLineBreakClass after, int next, bool atEnd,
            OfficeLineBreakClass afterNext) {
            if (after is OfficeLineBreakClass.PO or OfficeLineBreakClass.PR) {
                int position = b;
                if (before is OfficeLineBreakClass.CL or OfficeLineBreakClass.CP) position = Previous(position);
                while (position >= 0 && _classes[position] is OfficeLineBreakClass.SY or OfficeLineBreakClass.IS) position = Previous(position);
                if (position >= 0 && _classes[position] == OfficeLineBreakClass.NU &&
                    (before is OfficeLineBreakClass.CL or OfficeLineBreakClass.CP or OfficeLineBreakClass.SY or OfficeLineBreakClass.IS or OfficeLineBreakClass.NU))
                    return true;
            }
            if (before is OfficeLineBreakClass.PO or OfficeLineBreakClass.PR) {
                if (after == OfficeLineBreakClass.NU) return true;
                if (after == OfficeLineBreakClass.OP && !atEnd) {
                    if (afterNext == OfficeLineBreakClass.NU) return true;
                    if (afterNext == OfficeLineBreakClass.IS) {
                        int third = Next(next);
                        if (third < _classes.Length && _classes[third] == OfficeLineBreakClass.NU) return true;
                    }
                }
            }
            if (after == OfficeLineBreakClass.NU) {
                if (before is OfficeLineBreakClass.HY or OfficeLineBreakClass.IS) return true;
                int position = b;
                while (position >= 0 && _classes[position] is OfficeLineBreakClass.SY or OfficeLineBreakClass.IS) position = Previous(position);
                if (position >= 0 && _classes[position] == OfficeLineBreakClass.NU) return true;
            }
            return false;
        }
    }
}
