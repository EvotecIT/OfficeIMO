namespace AngleSharp.Js.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using System;

    /// <summary>
    /// Represents the WindowBase64 auxiliary interface implemented by Window.
    /// </summary>
    [DomName("WindowBase64")]
    [DomNoInterfaceObject]
    [DomExposed("Window")]
    [DomExposed("Worker")]
    public static class WindowBase64
    {
        #region Methods

        /// <summary>
        /// Takes the input data, in the form of a Unicode string containing
        /// only characters in the range U+0000 to U+00FF, each representing a
        /// binary byte with values 0x00 to 0xFF respectively, and converts it
        /// to its base64 representation, which it returns.
        /// </summary>
        /// <param name="window">The host.</param>
        /// <param name="data">String of bytes.</param>
        /// <returns>Base-64 representation.</returns>
        [DomName("btoa")]
        public static String Btoa(this IWindow window, String data)
        {
            var content = new Byte[data.Length];

            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] > 255)
                {
                    throw new DomException(DomError.InvalidCharacter);
                }

                content[i] = (Byte)data[i];
            }

            return Convert.ToBase64String(content);
        }

        /// <summary>
        /// Takes the input data, in the form of a Unicode string containing
        /// base64-encoded binary data, decodes it, and returns a string
        /// consisting of characters in the range U+0000 to U+00FF, each
        /// representing a binary byte with values 0x00 to 0xFF respectively,
        /// corresponding to that binary data.
        /// </summary>
        /// <param name="window">The host.</param>
        /// <param name="data">Base-64 representation.</param>
        /// <returns>String of bytes.</returns>
        [DomName("atob")]
        public static String Atob(this IWindow window, String data)
        {
            var chars = data.ToCharArray();
            var length = NormalizeCharacters(ref chars);
            var content = Convert.FromBase64CharArray(chars, 0, length);

            if (content.Length > chars.Length)
            {
                Array.Resize(ref chars, content.Length);
            }

            for (var i = 0; i < content.Length; i++)
            {
                chars[i] = (Char)content[i];
            }

            return new String(chars, 0, content.Length);
        }

        #endregion

        #region Helpers

        private static Int32 NormalizeCharacters(ref Char[] chars)
        {
            var length = chars.Length;
            ShiftCharacters(chars, ref length);
            var rem = length % 4;

            if (rem == 0)
            {
                while (length > 0 && chars[length - 1] == '=')
                {
                    length--;
                }

                rem = length % 4;
            }

            if (rem == 1)
            {
                throw new DomException(DomError.InvalidCharacter);
            }

            CheckCharacters(chars, length);

            if (rem != 0)
            {
                if (length + 4 - rem > chars.Length)
                {
                    Array.Resize(ref chars, length + 4 - rem);
                }

                while (rem++ != 4)
                {
                    chars[length++] = '=';
                }
            }

            return length;
        }

        private static void CheckCharacters(Char[] chars, Int32 length)
        {
            for (var i = 0; i < length; i++)
            {
                if (!IsLegalBase64Char(chars[i]))
                {
                    throw new DomException(DomError.InvalidCharacter);
                }
            }
        }

        private static void ShiftCharacters(Char[] chars, ref Int32 length)
        {
            var shift = 0;

            for (int i = 0; i < length; i++)
            {
                if (Char.IsWhiteSpace(chars[i]))
                {
                    shift++;
                }
                else if (shift > 0)
                {
                    var tmp = chars[i];
                    chars[i] = chars[i - shift];
                    chars[i - shift] = tmp;
                }
            }

            length -= shift;
        }

        private static Boolean IsLegalBase64Char(Char c) =>
            (c >= 'A' && c <= 'Z') ||
            (c >= 'a' && c <= 'z') ||
            (c >= '0' && c <= '9') ||
            (c == '+' || c == '/');

        #endregion
    }
}
