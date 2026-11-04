using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Encodes this document using the selected physical format and output encryption.</summary>
    public byte[] ToBytes(WordSaveOptions? options = null) {
        WordFileFormat format = options?.Format ?? WordFileFormat.Docx;
        string? password = ResolveSavePassword(options);
        if (password != null && format != WordFileFormat.Docx)
            throw new NotSupportedException("Password-protected DOC output is not supported.");
        byte[] bytes = ToWordBytes(format, options);
        return password == null ? bytes : OfficeEncryption.EncryptPackage(bytes, password);
    }
}
