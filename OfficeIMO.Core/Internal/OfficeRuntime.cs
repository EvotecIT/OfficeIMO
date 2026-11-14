namespace OfficeIMO.Core.Internal;

/// <summary>Runtime checks for package behavior that depends on the CLR rather than the compiled target.</summary>
internal static class OfficeRuntime {
    // Assembly.GetName allocates and parses the full name, and saves ask on every call.
    internal static readonly bool IsNetFramework = string.Equals(typeof(object).Assembly.GetName().Name,
        "mscorlib", StringComparison.Ordinal);
}
