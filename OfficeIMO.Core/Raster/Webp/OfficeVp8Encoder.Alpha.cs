// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private bool ComputeAlphaUsed(byte[] rgba, int width, int height, int stride) {
        var alphaOffset = 3;
        for (var y = 0; y < height; y++) {
            Checkpoint();
            var offset = y * stride + alphaOffset;
            for (var x = 0; x < width; x++) {
                if (rgba[offset] != 255) return true;
                offset += 4;
            }
        }
        return false;
    }

    private byte[] BuildAlphPayload(byte[] rgba, int width, int height, int stride) {
        var alpha = new byte[checked(width * height)];
        var dst = 0;
        for (var y = 0; y < height; y++) {
            Checkpoint();
            var offset = y * stride + 3;
            for (var x = 0; x < width; x++) {
                alpha[dst++] = rgba[offset];
                offset += 4;
            }
        }

        var filter = ChooseAlphaFilter(alpha, width, height);
        var payload = new byte[alpha.Length + 1];
        payload[0] = (byte)((filter & 0x3) << 2); // compression=0 (raw), filter, preprocessing=0
        if (filter == 0) {
            Buffer.BlockCopy(alpha, 0, payload, 1, alpha.Length);
            return payload;
        }

        ApplyAlphaFilterEncode(alpha, width, height, filter, payload, 1);
        return payload;
    }

    private int ChooseAlphaFilter(byte[] alpha, int width, int height) {
        var bestFilter = 0;
        var bestCost = ComputeAlphaFilterCost(alpha, width, height, filter: 0);
        for (var filter = 1; filter <= 3; filter++) {
            var cost = ComputeAlphaFilterCost(alpha, width, height, filter);
            if (cost < bestCost) {
                bestCost = cost;
                bestFilter = filter;
            }
        }

        return bestFilter;
    }

    private long ComputeAlphaFilterCost(byte[] alpha, int width, int height, int filter) {
        if (filter == 0) {
            long sum = 0;
            for (var i = 0; i < alpha.Length; i++) {
                sum += alpha[i];
            }
            return sum;
        }

        long cost = 0;
        for (var y = 0; y < height; y++) {
            Checkpoint();
            var row = y * width;
            for (var x = 0; x < width; x++) {
                var index = row + x;
                var predictor = GetAlphaPredictor(alpha, width, x, y, filter);
                var diff = alpha[index] - predictor;
                if (diff < 0) diff = -diff;
                cost += diff;
            }
        }

        return cost;
    }

    private void ApplyAlphaFilterEncode(byte[] alpha, int width, int height, int filter, byte[] output, int offset) {
        for (var y = 0; y < height; y++) {
            Checkpoint();
            var row = y * width;
            for (var x = 0; x < width; x++) {
                var index = row + x;
                var predictor = GetAlphaPredictor(alpha, width, x, y, filter);
                output[offset + index] = unchecked((byte)(alpha[index] - predictor));
            }
        }
    }

    private byte GetAlphaPredictor(byte[] alpha, int width, int x, int y, int filter) {
        var index = (y * width) + x;
        // WebP uses horizontal prediction for every filter on the first row and
        // vertical prediction on the first column, including filters 1 and 2.
        if (filter != 0 && y == 0) return x > 0 ? alpha[index - 1] : (byte)0;
        if (filter != 0 && x == 0) return alpha[index - width];
        var left = x > 0 ? alpha[index - 1] : (byte)0;
        var up = y > 0 ? alpha[index - width] : (byte)0;
        var upLeft = (x > 0 && y > 0) ? alpha[index - width - 1] : (byte)0;

        return filter switch {
            1 => left,
            2 => up,
            3 => ClampToByte(left + up - upLeft),
            _ => (byte)0
        };
    }

}
