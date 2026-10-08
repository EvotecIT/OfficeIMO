// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private bool EncodeBlockCoefficients(
        OfficeVp8BoolEncoder encoder,
        int[] probabilities,
        int blockType,
        int initialContext,
        int[] coefficientsNatural) {
        var prevContext = initialContext;
        var hasNonZero = false;

        bool previousWasZero = false;
        for (var coefficientIndex = blockType == OfficeVp8Tables.LumaAc ? 1 : 0; coefficientIndex < CoefficientsPerBlock; coefficientIndex++) {
            var band = CoeffBandTable[coefficientIndex];
            var naturalIndex = ZigZagToNaturalOrder[coefficientIndex];
            var value = coefficientsNatural[naturalIndex];
            var hasLater = HasNonZeroAfter(coefficientsNatural, coefficientIndex + 1);

            int token;
            int extraBits;
            if (value == 0) {
                token = hasLater ? 1 : 0;
                extraBits = 0;
            } else {
                token = GetTokenForMagnitude(Math.Abs(value), out extraBits);
            }

            WriteCoefficientToken(encoder, probabilities, blockType, band, prevContext, token, previousWasZero);

            if (token == 0) {
                break;
            }

            if (token > 1) {
                if (token >= 6) {
                    byte[] category = OfficeVp8Tables.CoefficientCategoryProbabilities[token - 6];
                    for (int bit = 0; bit < category.Length; bit++) {
                        encoder.WriteBool(category[bit], ((extraBits >> (category.Length - bit - 1)) & 1) != 0);
                    }
                }

                encoder.WriteBool(128, value < 0);
                hasNonZero = true;
            }

            prevContext = GetPrevContextAfter(token);
            previousWasZero = token == 1;
        }

        return hasNonZero;
    }

    private void WriteCoefficientToken(
        OfficeVp8BoolEncoder encoder,
        int[] probabilities,
        int blockType,
        int band,
        int prevContext,
        int token,
        bool skipEndOfBlock) {
        if ((uint)blockType >= CoeffBlockTypes) blockType = BlockTypeY;
        if ((uint)band >= CoeffBands) band = 0;
        if ((uint)prevContext >= CoeffPrevContexts) prevContext = 0;

        var node = skipEndOfBlock ? 2 : 0;
        while (true) {
            var probabilityIndex = node >> 1;
            var coeffIndex = GetCoeffIndex(blockType, band, prevContext, probabilityIndex);
            var probability = probabilities[coeffIndex];

            var left = CoeffTokenTree[node];
            var right = CoeffTokenTree[node + 1];

            if (ContainsToken(left, token)) {
                encoder.WriteBool(probability, false);
                if (left <= 0) return;
                node = left;
            } else {
                encoder.WriteBool(probability, true);
                if (right <= 0) return;
                node = right;
            }
        }
    }

    private int GetTokenForMagnitude(int magnitude, out int extraBits) {
        extraBits = 0;
        if (magnitude <= 1) return 2;
        if (magnitude == 2) return 3;
        if (magnitude == 3) return 4;
        if (magnitude == 4) return 5;

        for (var token = 6; token < CoeffTokenBaseMagnitude.Length; token++) {
            var baseMagnitude = CoeffTokenBaseMagnitude[token];
            var bitCount = CoeffTokenExtraBits[token];
            var max = baseMagnitude + ((1 << bitCount) - 1);
            if (magnitude <= max) {
                extraBits = magnitude - baseMagnitude;
                return token;
            }
        }

        extraBits = 0;
        return 11;
    }

    private int GetPrevContextAfter(int token) {
        return token switch {
            0 or 1 => 0,
            2 => 1,
            _ => 2,
        };
    }

    private bool HasNonZeroAfter(int[] coefficientsNatural, int zigZagStartIndex) {
        for (var i = zigZagStartIndex; i < CoefficientsPerBlock; i++) {
            var naturalIndex = ZigZagToNaturalOrder[i];
            if (coefficientsNatural[naturalIndex] != 0) return true;
        }
        return false;
    }

    private bool ContainsToken(int nodeValue, int token) {
        if (nodeValue <= 0) {
            return -nodeValue - 1 == token;
        }
        int left = CoeffTokenTree[nodeValue];
        int right = CoeffTokenTree[nodeValue + 1];
        return ContainsToken(left, token) || ContainsToken(right, token);
    }

}
