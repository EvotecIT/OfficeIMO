namespace OfficeIMO.Pdf;

/// <summary>Common PDF signature subfilters used by external signing workflows.</summary>
public enum PdfExternalSignatureSubFilter {
    /// <summary>Detached CMS/PKCS#7 signature, emitted as /adbe.pkcs7.detached.</summary>
    DetachedCms = 0,

    /// <summary>Detached CAdES signature, emitted as /ETSI.CAdES.detached.</summary>
    CadesDetached = 1,

    /// <summary>RFC 3161 document timestamp signature, emitted as /ETSI.RFC3161.</summary>
    DocumentTimestamp = 2
}
