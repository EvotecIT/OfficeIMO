using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Word;

/// <summary>Owns read-only file loads while preserving the lifetime of an explicit or enclosing document.</summary>
internal static class WordQuerySource {
    internal static void Visit(PSCmdlet command, string source, string[] path, string[] literalPath,
        WordDocument? document, Action<WordDocument> inspect) {
        if (source == "Path" || source == "LiteralPath") {
            bool literal = source == "LiteralPath";
            foreach (var file in OfficePath.ResolveInputs(command, literal ? literalPath : path, literal)) {
                using (var loaded = WordDocument.Load(file, new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly, PersistenceMode = OfficeIMO.DocumentPersistenceMode.Explicit })) inspect(loaded);
            }
            return;
        }
        inspect(document ?? OfficeDocumentContext.Find<WordDocument>() ?? WordDslContext.Current?.Document
            ?? throw new PSInvalidOperationException("Provide -Document or use this command inside a Word content block."));
    }
}
