using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static class Guard {
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNull<T>(T? value, string paramName) where T : class {
        if (value is null) throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNullOrEmpty(byte[]? value, string paramName) {
        if (value is null)
            throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");

        if (value.Length == 0)
            throw new System.ArgumentException($"Parameter '{paramName}' cannot be empty.", paramName);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void Positive(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0)
            throw new System.ArgumentOutOfRangeException(paramName, value, $"Parameter '{paramName}' must be a finite positive number.");
    }
}
