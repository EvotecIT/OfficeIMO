using System;

namespace OfficeIMO.Core.Internal;

/// <summary>Runtime checks for package behavior that depends on the CLR rather than the compiled target.</summary>
internal static class OfficeRuntime {
    internal static bool IsNetFramework => string.Equals(typeof(object).Assembly.GetName().Name,
        "mscorlib", StringComparison.Ordinal);
}
