using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The scanner finds the end of a plain scalar line with a table lookup over the first characters and a vectorized
// search after them. Both must stop where the character-class definition says, in block and flow context.
public class YamlScannerSearchTests {
    [Fact]
    public void PlainStopSearchMatchesTheCharacterClasses() {
        char[] alphabet = { 'a', 'Z', '0', ' ', '\t', '-', '.', ':', '#', '\r', '\n', '\0', '\u0001', '\u007F', ',', '[', ']', '{', '}', '\u00E9', '\u0085', '\u2028', '\uFEFF' };
        var random = new Random(20261110);
        var buffer = new char[48];
        for (int iteration = 0; iteration < 50_000; iteration++) {
            int length = random.Next(buffer.Length + 1);
            // Mostly plain text with an occasional special character, so stops land at every offset.
            for (int i = 0; i < length; i++) buffer[i] = random.Next(8) == 0 ? alphabet[random.Next(alphabet.Length)] : (char)('a' + random.Next(26));
            ReadOnlySpan<char> text = buffer.AsSpan(0, length);
            foreach (bool flow in new[] { false, true }) {
                Assert.Equal(Expected(text, flow), YamlChars.IndexOfPlainStop(text, flow));
            }
        }
    }

    private static int Expected(ReadOnlySpan<char> text, bool flow) {
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool control = (c < 0x20 && c != '\t' && c != '\n' && c != '\r') || c == '\u007F';
            if (control || c is ':' or '#' or '\r' or '\n' || (flow && c is ',' or '[' or ']' or '{' or '}')) return i;
        }
        return -1;
    }
}
