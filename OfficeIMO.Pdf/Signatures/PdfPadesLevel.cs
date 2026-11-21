namespace OfficeIMO.Pdf;

/// <summary>PAdES baseline level (ETSI EN 319 142-1) that validation established for a signature.</summary>
public enum PdfPadesLevel {
    /// <summary>No baseline level was established, or the signature was not cryptographically validated.</summary>
    None = 0,

    /// <summary>B-B: a valid ETSI.CAdES.detached signature carrying the ESS signing-certificate attribute.</summary>
    BaselineB = 1,

    /// <summary>B-T: B-B plus a valid signature time-stamp or a valid document time-stamp covering the signature.</summary>
    BaselineT = 2,

    /// <summary>B-LT: B-T plus Document Security Store data that settles revocation for every certificate checked.</summary>
    BaselineLT = 3,

    /// <summary>B-LTA: B-LT plus a valid document time-stamp covering the validation data used.</summary>
    BaselineLTA = 4
}
