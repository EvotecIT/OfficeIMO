namespace OfficeIMO.Pdf;

/// <summary>How an imposition operation handles a signed source PDF.</summary>
public enum PdfImpositionSignaturePolicy {
    /// <summary>Reject signed sources without altering them.</summary>
    Reject,
    /// <summary>Create an explicit unsigned source derivative before placing pages on new sheets.</summary>
    CreateUnsignedDerivative
}
