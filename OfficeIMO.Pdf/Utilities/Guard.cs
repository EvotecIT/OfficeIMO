using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static class Guard {
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void NotNull<T>(T? value, string paramName) where T : class {
        if (value is null) throw new System.ArgumentNullException(paramName, $"Parameter '{paramName}' cannot be null.");
    }
}
