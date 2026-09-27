// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    private static bool TryDecodeBlockCoefficients(
        OfficeVp8BoolDecoder decoder,
        OfficeVp8CoefficientProbabilities probabilities,
        int blockType,
        int initialContext,
        int dequantDc,
        int dequantAc,
        out int[] dequantizedCoefficients,
        out bool hasNonZero) {
        dequantizedCoefficients = new int[CoefficientsPerBlock];
        hasNonZero = false;

        var prevContext = initialContext;
        var coefficientIndex = blockType == OfficeVp8Tables.LumaAc ? 1 : 0;
        bool previousWasZero = false;

        while (coefficientIndex < CoefficientsPerBlock) {
            var band = CoeffBandTable[coefficientIndex];
            if (!TryReadCoefficientToken(
                decoder,
                probabilities,
                blockType,
                band,
                prevContext,
                out var tokenCode, previousWasZero)) {
                return false;
            }

            if (tokenCode == 0) {
                break;
            }

            var coeffValue = 0;
            if (tokenCode > 1) {
                if (!TryReadTokenExtraBits(decoder, tokenCode, out var extraBitsValue)) return false;
                var magnitude = ComputeTokenMagnitude(tokenCode, extraBitsValue);
                if (!TryReadSignedMagnitude(decoder, magnitude, out coeffValue)) return false;
            }

            if (coeffValue != 0) {
                hasNonZero = true;
            }

            var dequantFactor = coefficientIndex == 0 ? dequantDc : dequantAc;
            var naturalIndex = MapZigZagToNaturalIndex(coefficientIndex);
            dequantizedCoefficients[naturalIndex] = coeffValue * dequantFactor;

            var tokenInfo = ClassifyToken(tokenCode, band, prevContext);
            prevContext = tokenInfo.PrevContextAfter;
            previousWasZero = tokenCode == 1;
            coefficientIndex++;
        }

        return true;
    }

    private static bool TryReadTokenExtraBits(OfficeVp8BoolDecoder decoder, int tokenCode, out int extraBitsValue) {
        extraBitsValue = 0;
        if ((uint)tokenCode >= CoeffTokenExtraBits.Length) return false;

        var bitCount = CoeffTokenExtraBits[tokenCode];
        if (bitCount == 0) return true;
        foreach (byte probability in OfficeVp8Tables.CoefficientCategoryProbabilities[tokenCode - 6]) {
            if (!decoder.TryReadBool(probability, out bool bit)) return false;
            extraBitsValue = (extraBitsValue << 1) | (bit ? 1 : 0);
        }
        return true;
    }

    private static int ComputeTokenMagnitude(int tokenCode, int extraBitsValue) {
        if ((uint)tokenCode >= CoeffTokenBaseMagnitude.Length) return 0;
        if (tokenCode <= 1) return 0;
        if (tokenCode <= 5) return CoeffTokenBaseMagnitude[tokenCode];
        return CoeffTokenBaseMagnitude[tokenCode] + extraBitsValue;
    }

    private static bool TryReadSignedMagnitude(OfficeVp8BoolDecoder decoder, int magnitude, out int coefficientValue) {
        coefficientValue = magnitude;
        if (magnitude == 0) return true;
        if (!decoder.TryReadBool(128, out var signBit)) return false;
        coefficientValue = signBit ? -magnitude : magnitude;
        return true;
    }
}
