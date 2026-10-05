namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    // These immutable strings are shared by printable Latin tokens. Script selection,
    // substitutions, positioning, and logical clusters still use the normal shaping path.
    private static readonly string[] AsciiCharacters = CreateAsciiCharacters();

    private static string[] CreateAsciiCharacters() {
        var characters = new string[95];
        for (int index = 0; index < characters.Length; index++) characters[index] = ((char)(index + 32)).ToString();
        return characters;
    }

    private static bool IsPrintableAscii(string text, System.Threading.CancellationToken cancellationToken) {
        for (int index = 0; index < text.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if ((uint)(text[index] - 32) > 94U) return false;
        }
        return true;
    }
}
