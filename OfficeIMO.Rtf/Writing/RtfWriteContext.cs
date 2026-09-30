namespace OfficeIMO.Rtf.Writing;

/// <summary>Document defaults shared by every semantic story and table writer.</summary>
internal sealed class RtfWriteContext {
    internal RtfWriteContext(int? defaultLanguageId, int unicodeSkipCount, int? defaultParagraphStyleId, bool preserveStyleInheritance = false) {
        DefaultLanguageId = defaultLanguageId;
        UnicodeSkipCount = unicodeSkipCount;
        DefaultParagraphStyleId = defaultParagraphStyleId;
        PreserveStyleInheritance = preserveStyleInheritance || defaultParagraphStyleId.HasValue;
    }

    internal int? DefaultLanguageId { get; }
    internal int UnicodeSkipCount { get; }
    internal int? DefaultParagraphStyleId { get; }
    internal bool PreserveStyleInheritance { get; }

    internal RtfWriteContext ForCharacterScope(bool preserveStyleInheritance) =>
        PreserveStyleInheritance || !preserveStyleInheritance ? this :
        new RtfWriteContext(DefaultLanguageId, UnicodeSkipCount, DefaultParagraphStyleId, preserveStyleInheritance: true);
}
