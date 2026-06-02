namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a file embedded in a generated PDF and optionally associated with the document catalog.
/// </summary>
public sealed class PdfEmbeddedFile {
    private readonly byte[] _data;
    private string _fileName;
    private string? _mimeType;
    private string? _description;
    private PdfAssociatedFileRelationship _relationship;

    /// <summary>Creates a PDF embedded file description.</summary>
    public PdfEmbeddedFile(
        string fileName,
        byte[] data,
        string? mimeType = null,
        PdfAssociatedFileRelationship relationship = PdfAssociatedFileRelationship.Unspecified,
        string? description = null) {
        ValidateFileName(fileName, nameof(fileName));
        Guard.NotNullOrEmpty(data, nameof(data));
        ValidateOptionalMimeType(mimeType, nameof(mimeType));
        ValidateOptionalDescription(description, nameof(description));
        Guard.AssociatedFileRelationship(relationship, nameof(relationship));

        _fileName = fileName;
        _data = (byte[])data.Clone();
        _mimeType = mimeType;
        _description = description;
        _relationship = relationship;
    }

    /// <summary>Embedded file name as exposed in the PDF name tree.</summary>
    public string FileName {
        get => _fileName;
        set {
            ValidateFileName(value, nameof(FileName));
            _fileName = value;
        }
    }

    /// <summary>Embedded file bytes. The returned array is a defensive copy.</summary>
    public byte[] Data => (byte[])_data.Clone();

    internal byte[] DataSnapshot => (byte[])_data.Clone();

    /// <summary>Optional MIME type, for example "application/xml".</summary>
    public string? MimeType {
        get => _mimeType;
        set {
            ValidateOptionalMimeType(value, nameof(MimeType));
            _mimeType = value;
        }
    }

    /// <summary>Associated-file relationship emitted as /AFRelationship.</summary>
    public PdfAssociatedFileRelationship Relationship {
        get => _relationship;
        set {
            Guard.AssociatedFileRelationship(value, nameof(Relationship));
            _relationship = value;
        }
    }

    /// <summary>Optional human-readable file description.</summary>
    public string? Description {
        get => _description;
        set {
            ValidateOptionalDescription(value, nameof(Description));
            _description = value;
        }
    }

    internal PdfEmbeddedFile Clone() {
        return new PdfEmbeddedFile(FileName, _data, MimeType, Relationship, Description);
    }

    private static void ValidateFileName(string? value, string paramName) {
        Guard.NotNullOrWhiteSpace(value, paramName);
        for (int i = 0; i < value!.Length; i++) {
            char ch = value[i];
            if (ch == '/' || ch == '\\' || char.IsControl(ch) || System.Array.IndexOf(System.IO.Path.GetInvalidFileNameChars(), ch) >= 0) {
                throw new ArgumentException("PDF embedded file names must be simple file names without path separators, invalid file-name characters, or control characters.", paramName);
            }
        }
    }

    private static void ValidateOptionalMimeType(string? value, string paramName) {
        if (value == null) {
            return;
        }

        if (value.Length == 0 || string.IsNullOrWhiteSpace(value)) {
            throw new ArgumentException("PDF embedded file MIME type cannot be empty or whitespace.", paramName);
        }

        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsControl(ch) || char.IsWhiteSpace(ch)) {
                throw new ArgumentException("PDF embedded file MIME type cannot contain whitespace or control characters.", paramName);
            }
        }
    }

    private static void ValidateOptionalDescription(string? value, string paramName) {
        if (value == null) {
            return;
        }

        if (value.Length == 0) {
            throw new ArgumentException("PDF embedded file description cannot be empty.", paramName);
        }

        for (int i = 0; i < value.Length; i++) {
            if (char.IsControl(value[i])) {
                throw new ArgumentException("PDF embedded file description cannot contain control characters.", paramName);
            }
        }
    }
}
