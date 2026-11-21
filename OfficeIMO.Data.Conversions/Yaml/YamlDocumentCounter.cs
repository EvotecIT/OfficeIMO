namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Counts the documents of a UTF-8 YAML stream without parsing it. Document markers (<c>---</c> and <c>...</c> at the
/// start of a line) cannot appear inside content (YAML 1.2.2, c-forbidden), so a line scan is exact: each <c>---</c>
/// starts a document, and so does content before the first marker or after a <c>...</c>.
/// </summary>
internal static class YamlDocumentCounter {
    public static long Count(Stream stream) {
        byte[] buffer = ArrayPool<byte>.Shared.Rent(64 * 1024);
        try {
            long documents = 0;
            bool open = false;
            bool lineStart = true;
            var line = new LineState();
            int read;
            bool first = true;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
                int i = 0;
                if (first) {
                    first = false;
                    if (read >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF) i = 3;
                }

                for (; i < read; i++) {
                    byte b = buffer[i];
                    if (b == (byte)'\n' || b == (byte)'\r') {
                        Finish(ref line, ref documents, ref open);
                        lineStart = true;
                        continue;
                    }

                    if (lineStart) {
                        line = default;
                        lineStart = false;
                    }

                    line.Add(b);
                }
            }

            Finish(ref line, ref documents, ref open);
            return documents;
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    // Tracks only what classifies a line: its first non-space byte and whether it begins with a marker.
    private struct LineState {
        public int Length;
        public byte B0, B1, B2, B3;
        public byte FirstNonSpace;
        public bool HasContent;

        public void Add(byte b) {
            switch (Length) {
                case 0: B0 = b; break;
                case 1: B1 = b; break;
                case 2: B2 = b; break;
                case 3: B3 = b; break;
            }

            if (!HasContent && b != (byte)' ' && b != (byte)'\t') {
                HasContent = true;
                FirstNonSpace = b;
            }

            Length++;
        }

        private bool Marker(byte c) => Length >= 3 && B0 == c && B1 == c && B2 == c && (Length == 3 || B3 == (byte)' ' || B3 == (byte)'\t');

        public bool IsStart => Marker((byte)'-');
        public bool IsEnd => Marker((byte)'.');
        public bool IsDirective => Length > 0 && B0 == (byte)'%';
        public bool IsBlankOrComment => !HasContent || FirstNonSpace == (byte)'#';
    }

    private static void Finish(ref LineState line, ref long documents, ref bool open) {
        if (line.Length == 0) return;
        if (line.IsStart) {
            documents++;
            open = true;
        } else if (line.IsEnd) {
            open = false;
        } else if (!open && !line.IsDirective && !line.IsBlankOrComment) {
            documents++;
            open = true;
        }

        line = default;
    }
}
