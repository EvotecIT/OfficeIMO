using System;

namespace OfficeIMO.Internal {
    /// <summary>Cross-target enum enumeration that remains statically analyzable for NativeAOT.</summary>
    internal static class EnumCompat {
        internal static TEnum[] GetValues<TEnum>() where TEnum : struct, Enum {
#if NET5_0_OR_GREATER
            return Enum.GetValues<TEnum>();
#else
            return (TEnum[])Enum.GetValues(typeof(TEnum));
#endif
        }
    }
}
