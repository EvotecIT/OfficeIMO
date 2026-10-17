namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    // Only bitmap MQ statistics are retained. Arithmetic integer and symbol-ID coders start afresh per dictionary.
    private sealed class SymbolBitmapContexts {
        internal readonly int Parameters;
        internal readonly sbyte[] Adaptive, RefinementAdaptive;
        internal readonly byte[] Generic, Refinement;

        internal SymbolBitmapContexts(int flags, sbyte[] adaptive, sbyte[] refinementAdaptive, byte[] generic, byte[] refinement) {
            Parameters = flags & 0x1C03; // SDHUFF, SDREFAGG, SDTEMPLATE and SDRTEMPLATE (7.4.2.2).
            Adaptive = adaptive;
            RefinementAdaptive = refinementAdaptive;
            Generic = generic;
            Refinement = refinement;
        }
    }

    private sealed partial class DecodingContext {
        private SymbolBitmapContexts CreateSymbolBitmapContexts(Segment segment, int flags, sbyte[] adaptive, sbyte[] refinementAdaptive) {
            byte[] generic, refinement;
            if ((flags & 0x100) != 0) {
                // It is the last referred symbol dictionary, not the last dictionary with retained statistics.
                SymbolBitmapContexts? retained = null;
                foreach (int referred in segment.Referred) {
                    Charge(1);
                    if (_symbols.ContainsKey(referred)) _retainedSymbolContexts.TryGetValue(referred, out retained);
                }
                if (retained == null)
                    throw new InvalidDataException("JBIG2 bitmap contexts require a retained referred symbol dictionary.");
                if (retained.Parameters != (flags & 0x1C03) || !retained.Adaptive.SequenceEqual(adaptive) ||
                    !retained.RefinementAdaptive.SequenceEqual(refinementAdaptive))
                    throw new InvalidDataException("JBIG2 retained bitmap context parameters do not match the symbol dictionary.");
                // Copy before decoding so two dictionaries can independently reuse the same retained state.
                generic = CopyContexts(retained.Generic);
                refinement = CopyContexts(retained.Refinement);
            } else {
                generic = CreateContexts(1 << 16);
                refinement = CreateContexts(1 << 14);
            }
            return new SymbolBitmapContexts(flags, adaptive, refinementAdaptive, generic, refinement);
        }

        private byte[] CopyContexts(byte[] source) {
            byte[] result = CreateContexts(source.Length);
            Charge(source.Length);
            Buffer.BlockCopy(source, 0, result, 0, source.Length);
            return result;
        }
    }
}
