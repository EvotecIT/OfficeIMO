namespace OfficeIMO.Provenance;

/// <summary>
/// Recognizes generative AI products named in generator fields such as XMP <c>xmp:CreatorTool</c> or EXIF <c>Software</c>.
/// A match means the producer declared the tool; it is not detection of how the content was made.
/// </summary>
internal static class OfficeProvenanceAiGenerators {
    // Product names only, matched case-insensitively as whole words, so "Firefly" alone or "Gemini" in an unrelated
    // product name does not match. Extend this list when a generator documents the name it writes.
    private static readonly string[] Names = {
        "Midjourney",
        "DALL-E", "DALL·E", "DALLE",
        "ChatGPT", "OpenAI",
        "Adobe Firefly",
        "Stable Diffusion", "StableDiffusion", "AUTOMATIC1111", "ComfyUI", "InvokeAI", "NovelAI", "Fooocus",
        "Google Gemini", "Google Imagen",
        "Bing Image Creator", "Microsoft Designer",
        "Leonardo.Ai", "Ideogram", "FLUX.1", "Recraft"
    };

    /// <summary>Whether <paramref name="value"/> names a generative AI product.</summary>
    internal static bool IsAiGenerator(string? value) => TryMatch(value, out _);

    /// <summary>
    /// Finds the generative AI product that <paramref name="value"/> names and returns it as listed here, so reports carry the product
    /// rather than the rest of the field, which can hold unrelated or private text.
    /// </summary>
    internal static bool TryMatch(string? value, out string product) {
        product = string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return false;
        foreach (string name in Names) {
            int start = 0;
            while (start <= value!.Length - name.Length) {
                int index = value.IndexOf(name, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0) break;
                int end = index + name.Length;
                if ((index == 0 || !char.IsLetterOrDigit(value[index - 1])) &&
                    (end == value.Length || !char.IsLetterOrDigit(value[end]))) {
                    product = name;
                    return true;
                }
                start = index + 1;
            }
        }
        return false;
    }

    /// <summary>
    /// Whether <paramref name="text"/> is a generation-settings record in the "Steps: …, Sampler: …, Seed: …" form that Stable
    /// Diffusion web UIs write into EXIF <c>UserComment</c>, or names an AI product.
    /// </summary>
    internal static bool IsAiGenerationRecord(string? text) =>
        text != null &&
        (text.IndexOf("Steps: ", StringComparison.Ordinal) >= 0 && text.IndexOf("Seed: ", StringComparison.Ordinal) >= 0 &&
            (text.IndexOf("Sampler: ", StringComparison.Ordinal) >= 0 || text.IndexOf("CFG scale: ", StringComparison.Ordinal) >= 0) ||
            IsAiGenerator(text));

    /// <summary>
    /// Whether an EXIF <c>UserComment</c> value is an AI generation record. A UTF-16 comment without a byte-order mark is read in
    /// both byte orders, because no byte count can decide the order for every text.
    /// </summary>
    internal static bool IsAiUserComment(byte[] data, int offset, int length) {
        if (IsAiGenerationRecord(DecodeUserComment(data, offset, length))) return true;
        if (length < 10 || !OfficeProvenanceBinary.MatchesAscii(data, offset, "UNICODE\0")) return false;
        int textOffset = offset + 8;
        int textLength = (length - 8) & ~1;
        if (data[textOffset] == 0xFE && data[textOffset + 1] == 0xFF || data[textOffset] == 0xFF && data[textOffset + 1] == 0xFE) return false;
        return IsAiGenerationRecord(System.Text.Encoding.BigEndianUnicode.GetString(data, textOffset, textLength)) ||
            IsAiGenerationRecord(System.Text.Encoding.Unicode.GetString(data, textOffset, textLength));
    }

    /// <summary>
    /// Decodes an EXIF <c>UserComment</c> value: an eight-byte character-code prefix followed by the text. ASCII, Unicode (UTF-16 in
    /// either byte order) and undefined (read as UTF-8) codes are supported; JIS and unknown codes return null.
    /// </summary>
    internal static string? DecodeUserComment(byte[] data, int offset, int length) {
        if (length < 8 || offset < 0 || offset > data.Length - length) return null;
        int textOffset = offset + 8;
        int textLength = length - 8;
        if (OfficeProvenanceBinary.MatchesAscii(data, offset, "ASCII\0\0\0")) {
            return System.Text.Encoding.ASCII.GetString(data, textOffset, textLength).TrimEnd('\0');
        }
        if (OfficeProvenanceBinary.MatchesAscii(data, offset, "UNICODE\0")) {
            if (textLength >= 2 && (data[textOffset] == 0xFE && data[textOffset + 1] == 0xFF || data[textOffset] == 0xFF && data[textOffset + 1] == 0xFE)) {
                System.Text.Encoding marked = data[textOffset] == 0xFE ? System.Text.Encoding.BigEndianUnicode : System.Text.Encoding.Unicode;
                int markedLength = textLength - 2;
                return marked.GetString(data, textOffset + 2, markedLength - markedLength % 2).TrimEnd('\0');
            }
            // Without a byte-order mark, the byte order with more zero high bytes is the one the text was written in.
            int evenZeros = 0, oddZeros = 0;
            for (int index = 0; index + 1 < textLength; index += 2) {
                if (data[textOffset + index] == 0) evenZeros++;
                if (data[textOffset + index + 1] == 0) oddZeros++;
            }
            System.Text.Encoding encoding = evenZeros > oddZeros ? System.Text.Encoding.BigEndianUnicode : System.Text.Encoding.Unicode;
            return encoding.GetString(data, textOffset, textLength - textLength % 2).TrimEnd('\0');
        }
        for (int index = 0; index < 8; index++) {
            if (data[offset + index] != 0) return null;
        }
        return System.Text.Encoding.UTF8.GetString(data, textOffset, textLength).TrimEnd('\0');
    }
}
