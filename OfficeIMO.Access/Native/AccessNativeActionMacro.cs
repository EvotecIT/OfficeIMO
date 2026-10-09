using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    /// <summary>Qualified legacy single StopMacro envelope. Other action/embedded/data representations remain opaque.</summary>
    internal static class AccessNativeActionMacro {
        internal static AccessActionMacroInfo? Read(byte[] bytes) {
            if (bytes.Length != 76 || U32(bytes, 0) != 0 || U16(bytes, 12) != 0 || U16(bytes, 14) != 2 || U16(bytes, 32) != 4
                || System.Text.Encoding.Unicode.GetString(bytes, 34, 6) != "33\0" || U16(bytes, 40) != 45 || U16(bytes, 42) != 1 || U32(bytes, 72) != 0) return null;
            // The qualified envelope has no conditions, parameters, submacros, embedded actions or trailing data.
            for (int i = 4; i < 72; i++) if (!(i >= 12 && i < 16 || i >= 32 && i < 44) && bytes[i] != 255) return null;
            return new AccessActionMacroInfo(new[] { "StopMacro" });
        }
    }
}
