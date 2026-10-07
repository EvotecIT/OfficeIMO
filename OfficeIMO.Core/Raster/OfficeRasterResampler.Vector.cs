#if NET8_0_OR_GREATER
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    // Four independent channels share a tap but retain double precision and tap order.
    // Separate multiply/add intrinsics preserve the scalar rounding sequence; alpha
    // receives its original weight while RGB retains the original premultiply formula.
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static void AccumulateBytesVector(byte[] input, int baseOffset, int stride,
        AxisContributions contributions, int destination, float[] output, int target, CancellationToken cancellationToken) {
        int start = contributions.Starts[destination];
        int count = contributions.Counts[destination];
        int weightOffset = contributions.Offsets[destination];
        double[] weights = contributions.Weights;
        Vector256<double> sum = Vector256<double>.Zero;
        for (int index = 0; index < count; index++) {
            if ((index & 1023) == 1023) cancellationToken.ThrowIfCancellationRequested();
            int source = baseOffset + ((start + index) * stride);
            // Contribution ranges guarantee a complete RGBA pixel at this offset.
            uint packed = Unsafe.ReadUnaligned<uint>(ref input[source]);
            Vector128<int> channels = Sse41.ConvertToVector128Int32(Vector128.CreateScalar(packed).AsByte());
            double weight = weights[weightOffset + index];
            double premultiply = weight * (packed >> 24) / 255D;
            Vector256<double> factors = Avx.Blend(Vector256.Create(premultiply), Vector256.Create(weight), 8);
            sum = Avx.Add(sum, Avx.Multiply(Avx.ConvertToVector256Double(channels), factors));
        }
        Avx.ConvertToVector128Single(sum).StoreUnsafe(ref MemoryMarshal.GetArrayDataReference(output), (nuint)target);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static void AccumulateFloatsVector(float[] input, int baseOffset, int stride,
        AxisContributions contributions, int destination, byte[] output, int target,
        OfficeRasterResamplingColorSpace colorSpace, CancellationToken cancellationToken) {
        int start = contributions.Starts[destination];
        int count = contributions.Counts[destination];
        int weightOffset = contributions.Offsets[destination];
        double[] weights = contributions.Weights;
        Vector256<double> sum = Vector256<double>.Zero;
        for (int index = 0; index < count; index++) {
            if ((index & 1023) == 1023) cancellationToken.ThrowIfCancellationRequested();
            int source = baseOffset + ((start + index) * stride);
            Vector128<float> channels = Vector128.LoadUnsafe(ref MemoryMarshal.GetArrayDataReference(input), (nuint)source);
            Vector256<double> weight = Vector256.Create(weights[weightOffset + index]);
            sum = Avx.Add(sum, Avx.Multiply(Avx.ConvertToVector256Double(channels), weight));
        }
        WriteStraightRgba(output, target, sum.GetElement(0), sum.GetElement(1), sum.GetElement(2), sum.GetElement(3), colorSpace);
    }
}
#endif
