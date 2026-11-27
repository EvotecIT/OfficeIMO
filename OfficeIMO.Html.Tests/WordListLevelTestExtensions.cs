using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

internal static class WordListLevelTestExtensions {
    /// <summary>The SDK element of a level in a list, for tests that build numbering the public API doesn't cover.</summary>
    internal static Level OpenXmlLevel(this WordListLevel level) => (Level)level.Node.SdkElement!;
}
