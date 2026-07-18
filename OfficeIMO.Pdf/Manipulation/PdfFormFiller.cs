namespace OfficeIMO.Pdf;

/// <summary>
/// Updates simple AcroForm field values in parser-supported PDFs.
/// </summary>
internal static partial class PdfFormFiller {
    private const string UnsupportedFlattenWidgetMessage = "Only simple text, choice, and button AcroForm widgets with rectangles are supported for flattening by OfficeIMO.Pdf yet.";
    private const string UnsupportedFlattenAnnotationMessage = "Only simple text, choice, and button AcroForm widgets referenced from page annotations are supported for flattening by OfficeIMO.Pdf yet.";
    private const int MultilineFlag = 4096;
    private const int PasswordFlag = 8192;
    private const int RadioButtonFlag = 32768;
    private const int EditableChoiceFlag = 262144;
    private const int MultiSelectChoiceFlag = 2097152;
    private const int CombFlag = 16777216;
}
