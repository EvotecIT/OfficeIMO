using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party text stamping helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static partial class PdfStamper {
    private const int FontPseudoObjectNumber = -1;
    private const int ImagePseudoObjectNumber = -2;
    private const int ImageSoftMaskPseudoObjectNumber = -3;

}
