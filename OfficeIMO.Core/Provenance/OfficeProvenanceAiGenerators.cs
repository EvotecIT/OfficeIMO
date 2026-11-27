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
    internal static bool IsAiGenerator(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        foreach (string name in Names) {
            int start = 0;
            while (start <= value!.Length - name.Length) {
                int index = value.IndexOf(name, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0) break;
                int end = index + name.Length;
                if ((index == 0 || !char.IsLetterOrDigit(value[index - 1])) &&
                    (end == value.Length || !char.IsLetterOrDigit(value[end]))) return true;
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
