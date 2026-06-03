namespace OfficeIMO.Pdf;

/// <summary>
/// Describes Factur-X/ZUGFeRD XMP extension metadata for a generated e-invoice PDF.
/// </summary>
public sealed class PdfElectronicInvoiceMetadata {
    /// <summary>Factur-X/ZUGFeRD 2.1+ XMP namespace URI.</summary>
    public const string FacturXNamespaceUri = "urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#";

    private string _documentType;
    private string _documentFileName;
    private string _version;
    private string _conformanceLevel;

    /// <summary>Creates e-invoice XMP metadata.</summary>
    public PdfElectronicInvoiceMetadata(string documentType, string documentFileName, string version, string conformanceLevel) {
        ValidateToken(documentType, nameof(documentType), "PDF e-invoice document type cannot be empty.");
        ValidateFileName(documentFileName, nameof(documentFileName));
        ValidateToken(version, nameof(version), "PDF e-invoice version cannot be empty.");
        ValidateToken(conformanceLevel, nameof(conformanceLevel), "PDF e-invoice conformance level cannot be empty.");

        _documentType = documentType.Trim();
        _documentFileName = documentFileName.Trim();
        _version = version.Trim();
        _conformanceLevel = conformanceLevel.Trim();
    }

    /// <summary>Document type declared in XMP metadata, normally <c>INVOICE</c>.</summary>
    public string DocumentType {
        get => _documentType;
        set {
            ValidateToken(value, nameof(DocumentType), "PDF e-invoice document type cannot be empty.");
            _documentType = value.Trim();
        }
    }

    /// <summary>Embedded XML invoice file name declared in XMP metadata.</summary>
    public string DocumentFileName {
        get => _documentFileName;
        set {
            ValidateFileName(value, nameof(DocumentFileName));
            _documentFileName = value.Trim();
        }
    }

    /// <summary>Factur-X/ZUGFeRD XMP schema version value.</summary>
    public string Version {
        get => _version;
        set {
            ValidateToken(value, nameof(Version), "PDF e-invoice version cannot be empty.");
            _version = value.Trim();
        }
    }

    /// <summary>Profile/conformance level declared in XMP metadata, for example <c>EN 16931</c>.</summary>
    public string ConformanceLevel {
        get => _conformanceLevel;
        set {
            ValidateToken(value, nameof(ConformanceLevel), "PDF e-invoice conformance level cannot be empty.");
            _conformanceLevel = value.Trim();
        }
    }

    /// <summary>Creates Factur-X/ZUGFeRD 2.1+ metadata with the canonical XML attachment file name.</summary>
    public static PdfElectronicInvoiceMetadata FacturX(string conformanceLevel = "EN 16931", string version = "1.0") {
        return new PdfElectronicInvoiceMetadata("INVOICE", "factur-x.xml", version, conformanceLevel);
    }

    internal PdfElectronicInvoiceMetadata Clone() {
        return new PdfElectronicInvoiceMetadata(DocumentType, DocumentFileName, Version, ConformanceLevel);
    }

    private static void ValidateToken(string? value, string paramName, string message) {
        if (string.IsNullOrWhiteSpace(value)) {
            throw new ArgumentException(message, paramName);
        }

        for (int i = 0; i < value!.Length; i++) {
            if (char.IsControl(value[i])) {
                throw new ArgumentException(message, paramName);
            }
        }
    }

    private static void ValidateFileName(string? value, string paramName) {
        ValidateToken(value, paramName, "PDF e-invoice document file name cannot be empty.");
        string fileName = value!.Trim();
        for (int i = 0; i < fileName.Length; i++) {
            char ch = fileName[i];
            if (ch == '/' || ch == '\\' || char.IsControl(ch) || Array.IndexOf(Path.GetInvalidFileNameChars(), ch) >= 0) {
                throw new ArgumentException("PDF e-invoice document file name must be a simple file name.", paramName);
            }
        }
    }
}
