using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static int FilterKernelLength(double sigma) => sigma <= 0D ? 1 : 2 * (int)Math.Ceiling(3D * sigma) + 1;

    private static float[] BlurFilterPixels(float[] input, int width, int height, double sx, double sy, CancellationToken token) {
        float[] horizontal = ConvolveFilterPixels(input, width, height, sx, true, token);
        return ConvolveFilterPixels(horizontal, width, height, sy, false, token);
    }

    private static float[] ConvolveFilterPixels(float[] input, int width, int height, double sigma, bool horizontal, CancellationToken token) {
        var output = new float[input.Length];
        if (sigma <= 0D) { Array.Copy(input, output, input.Length); return output; }
        int radius = (FilterKernelLength(sigma) - 1) / 2;
        var weights = new double[2 * radius + 1];
        double sum = 0D;
        for (int k = -radius; k <= radius; k++) {
            double distance = k / sigma;
            double weight = Math.Exp(-0.5D * distance * distance);
            weights[k + radius] = weight;
            sum += weight;
        }
        for (int k = 0; k < weights.Length; k++) weights[k] /= sum;
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                int dest = (y * width + x) * 4;
                for (int k = -radius; k <= radius; k++) {
                    int xx = horizontal ? x + k : x, yy = horizontal ? y : y + k;
                    if (xx < 0 || xx >= width || yy < 0 || yy >= height) continue;
                    int src = (yy * width + xx) * 4;
                    double weight = weights[k + radius];
                    for (int c = 0; c < 4; c++) output[dest + c] += (float)(input[src + c] * weight);
                }
            }
        }
        return output;
    }

    private static float[] OffsetFilterPixels(float[] input, int width, int height, double dx, double dy, CancellationToken token) {
        var output = new float[input.Length];
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                double px = x - dx, py = y - dy;
                if (px <= -1D || px >= width || py <= -1D || py >= height) continue;
                int xx = (int)Math.Floor(px), yy = (int)Math.Floor(py);
                double fx = px - xx, fy = py - yy;
                int dest = (y * width + x) * 4;
                for (int iy = 0; iy <= 1; iy++) {
                    for (int ix = 0; ix <= 1; ix++) {
                        int sx = xx + ix, sy = yy + iy;
                        if (sx < 0 || sx >= width || sy < 0 || sy >= height) continue;
                        double weight = (ix == 0 ? 1D - fx : fx) * (iy == 0 ? 1D - fy : fy);
                        int src = (sy * width + sx) * 4;
                        for (int c = 0; c < 4; c++) output[dest + c] += (float)(input[src + c] * weight);
                    }
                }
            }
        }
        return output;
    }

    private static float[] MatrixFilterPixels(float[] input, int width, int height, double[] matrix, CancellationToken token) {
        var output = new float[input.Length];
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int i = y * width * 4; i < (y + 1) * width * 4; i += 4) {
                double a = input[i + 3];
                double r = a > 0D ? input[i] / a : 0D, g = a > 0D ? input[i + 1] / a : 0D, b = a > 0D ? input[i + 2] / a : 0D;
                double alpha = FilterClamp(matrix[15] * r + matrix[16] * g + matrix[17] * b + matrix[18] * a + matrix[19]);
                for (int c = 0; c < 3; c++) {
                    int row = c * 5;
                    output[i + c] = (float)(FilterClamp(matrix[row] * r + matrix[row + 1] * g + matrix[row + 2] * b + matrix[row + 3] * a + matrix[row + 4]) * alpha);
                }
                output[i + 3] = (float)alpha;
            }
        }
        return output;
    }

    private static float[] CompositeFilterPixels(float[] source, float[] backdrop, int width, int height, OfficeBlendMode mode, CancellationToken token) {
        var output = new float[source.Length];
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int i = y * width * 4; i < (y + 1) * width * 4; i += 4) {
                double sa = source[i + 3], ba = backdrop[i + 3];
                for (int c = 0; c < 3; c++) {
                    double cs = sa > 0D ? source[i + c] / sa : 0D, cb = ba > 0D ? backdrop[i + c] / ba : 0D;
                    double blend = OfficeRasterCanvas.BlendComponent(cb, cs, mode);
                    output[i + c] = (float)((1D - ba) * source[i + c] + (1D - sa) * backdrop[i + c] + sa * ba * blend);
                }
                output[i + 3] = (float)(sa + ba * (1D - sa));
            }
        }
        return output;
    }
}
