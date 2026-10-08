// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private void ConvertRgbaToYuv420(
        byte[] rgba,
        int width,
        int height,
        int stride,
        out byte[] yPlane,
        out byte[] uPlane,
        out byte[] vPlane) {
        var chromaWidth = (width + 1) >> 1;
        var chromaHeight = (height + 1) >> 1;

        yPlane = new byte[checked(width * height)];
        uPlane = new byte[checked(chromaWidth * chromaHeight)];
        vPlane = new byte[checked(chromaWidth * chromaHeight)];

        for (int cy = 0; cy < chromaHeight; cy++) {
            Checkpoint();
            for (int cx = 0; cx < chromaWidth; cx++) {
                int sumU = 0, sumV = 0, count = 0;
                for (int dy = 0; dy < 2; dy++) {
                    int y = cy * 2 + dy;
                    if (y >= height) break;
                    for (int dx = 0; dx < 2; dx++) {
                        int x = cx * 2 + dx;
                        if (x >= width) break;
                        int at = y * stride + x * 4;
                        int r = rgba[at], g = rgba[at + 1], b = rgba[at + 2];
                        yPlane[y * width + x] = ClampToByte(16 + ((16839 * r + 33059 * g + 6420 * b + 32768) >> 16));
                        sumU += 128 + ((-9719 * r - 19081 * g + 28800 * b + 32768) >> 16);
                        sumV += 128 + ((28800 * r - 24116 * g - 4684 * b + 32768) >> 16);
                        count++;
                    }
                }
                int chromaAt = cy * chromaWidth + cx;
                uPlane[chromaAt] = ClampToByte((sumU + (count >> 1)) / count);
                vPlane[chromaAt] = ClampToByte((sumV + (count >> 1)) / count);
            }
        }
    }

}
