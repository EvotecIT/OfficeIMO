namespace OfficeIMO.Epub;

/// <summary>Describes the structural EPUB signature carrier without asserting cryptographic validity.</summary>
public sealed class EpubSignatureInfo {
    internal static readonly EpubSignatureInfo NotPresent = new(false, true, 0);

    internal EpubSignatureInfo(bool isPresent, bool isWellFormed, int xmlSignatureCount) {
        IsPresent = isPresent;
        IsWellFormed = isWellFormed;
        XmlSignatureCount = xmlSignatureCount;
    }

    /// <summary>Whether META-INF/signatures.xml exists.</summary>
    public bool IsPresent { get; }

    /// <summary>Whether the signature carrier was within limits and parsed as well-formed XML.</summary>
    public bool IsWellFormed { get; }

    /// <summary>Number of XML Digital Signature elements declared by the carrier.</summary>
    public int XmlSignatureCount { get; }
}
