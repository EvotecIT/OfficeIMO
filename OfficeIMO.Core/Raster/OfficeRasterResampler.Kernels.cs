using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    private static double KernelRadius(OfficeRasterResamplingMode mode) => mode switch {
        OfficeRasterResamplingMode.Box => .5D,
        OfficeRasterResamplingMode.Triangle or OfficeRasterResamplingMode.Hermite => 1D,
        OfficeRasterResamplingMode.Lanczos3 or OfficeRasterResamplingMode.Welch => 3D,
        OfficeRasterResamplingMode.Lanczos5 => 5D,
        OfficeRasterResamplingMode.Lanczos8 => 8D,
        _ => 2D
    };

    private static double ReconstructionKernel(double value, OfficeRasterResamplingMode mode) {
        double x = Math.Abs(value);
        return mode switch {
            OfficeRasterResamplingMode.Box => x <= .5D ? 1D : 0D,
            OfficeRasterResamplingMode.Triangle => Math.Max(0D, 1D - x),
            OfficeRasterResamplingMode.Hermite => x < 1D ? (2D * x - 3D) * x * x + 1D : 0D,
            OfficeRasterResamplingMode.Lanczos2 or OfficeRasterResamplingMode.Lanczos3 or
                OfficeRasterResamplingMode.Lanczos5 or OfficeRasterResamplingMode.Lanczos8 =>
                x < KernelRadius(mode) ? Sinc(value) * Sinc(value / KernelRadius(mode)) : 0D,
            OfficeRasterResamplingMode.Welch => x < 3D ? Sinc(value) * (1D - x * x / 9D) : 0D,
            OfficeRasterResamplingMode.MitchellNetravali => CubicBc(x, 1D / 3D, 1D / 3D),
            OfficeRasterResamplingMode.Robidoux => CubicBc(x, .37821575509399867D, .31089212245300067D),
            OfficeRasterResamplingMode.RobidouxSharp => CubicBc(x, .2620145123990142D, .3689927438004929D),
            OfficeRasterResamplingMode.Spline => CubicBc(x, 1D, 0D),
            OfficeRasterResamplingMode.Bicubic or OfficeRasterResamplingMode.CatmullRom => CubicBc(x, 0D, .5D),
            _ => throw new ArgumentOutOfRangeException(nameof(mode))
        };
    }

    private static double Sinc(double x) => Math.Abs(x) < 1E-12D ? 1D : Math.Sin(Math.PI * x) / (Math.PI * x);

    // Mitchell-Netravali's two-parameter cubic family; all kernels share planning,
    // antialiasing, premultiplied-alpha filtering and retained-memory accounting.
    private static double CubicBc(double x, double b, double c) {
        if (x < 1D) return ((12D - 9D * b - 6D * c) * x * x * x +
            (-18D + 12D * b + 6D * c) * x * x + 6D - 2D * b) / 6D;
        if (x < 2D) return ((-b - 6D * c) * x * x * x +
            (6D * b + 30D * c) * x * x + (-12D * b - 48D * c) * x + 8D * b + 24D * c) / 6D;
        return 0D;
    }
}
