using System;

namespace OfficeIMO.Security {
    /// <summary>Provides allocation-light encoded-length guards for bounded OPC signature base64 values.</summary>
    internal static class OfficePackageBase64 {
        /// <summary>
        /// Returns whether non-whitespace base64 characters can exceed the decoded byte limit.
        /// The accepted whitespace set matches <see cref="Convert.FromBase64String(string)"/>.
        /// </summary>
        internal static bool ExceedsDecodedByteLimit(string value, long maxDecodedBytes) {
            long maxEncodedCharacters = GetMaxEncodedCharacters(maxDecodedBytes);
            long encodedCharacters = 0;
            foreach (char character in value) {
                if (character == ' ' || character == '\t' || character == '\r' || character == '\n') {
                    continue;
                }
                encodedCharacters++;
                if (encodedCharacters > maxEncodedCharacters) return true;
            }
            return false;
        }

        private static long GetMaxEncodedCharacters(long maxDecodedBytes) =>
            maxDecodedBytes > (long.MaxValue / 4L) * 3L
                ? long.MaxValue
                : ((maxDecodedBytes + 2L) / 3L) * 4L;
    }
}
