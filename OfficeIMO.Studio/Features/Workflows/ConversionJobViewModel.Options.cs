using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionJobViewModel {
    public bool SupportsPageSelection => Route.Route.SupportsPageSelection;
    public bool SupportsPdfCompression => Route.Route.SupportsPdfCompression;
    public bool IsWordImport => Route.Route.Id == "pdf-docx";
    public bool IsPowerPointImport => Route.Route.Id == "pdf-pptx";
    public bool IsWorksheetExport => Route.Route.Id == "xlsx-pdf";
    public bool IsHtmlExport => Route.Route.Id == "pdf-html";
    public bool CanEditOptions => State == ConversionJobState.Queued || CanRetry;
    public bool UsesRasterPages => IsWordImport && WordMode == PdfWordImportMode.VisualPages ||
        IsPowerPointImport && PowerPointMode is PdfPowerPointImportMode.VisualPages or PdfPowerPointImportMode.HybridVisualAndEditableTables;
    public IReadOnlyList<PdfWordImportMode> WordModes { get; } = Enum.GetValues<PdfWordImportMode>();
    public IReadOnlyList<PdfPowerPointImportMode> PowerPointModes { get; } = [
        PdfPowerPointImportMode.EditableContent, PdfPowerPointImportMode.VisualPages,
        PdfPowerPointImportMode.HybridVisualAndEditableTables, PdfPowerPointImportMode.EditableTables];
    public IReadOnlyList<ExcelPdfWorksheetLayoutMode> WorksheetLayouts { get; } = Enum.GetValues<ExcelPdfWorksheetLayoutMode>();
    public IReadOnlyList<PdfHtmlProfile> HtmlProfiles { get; } = Enum.GetValues<PdfHtmlProfile>();

    [ObservableProperty] private string _pageRanges = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(UsesRasterPages)), NotifyPropertyChangedFor(nameof(Fidelity))]
    private PdfWordImportMode _wordMode = PdfWordImportMode.EditableContent;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(UsesRasterPages)), NotifyPropertyChangedFor(nameof(Fidelity))]
    private PdfPowerPointImportMode _powerPointMode = PdfPowerPointImportMode.EditableContent;
    [ObservableProperty] private ExcelPdfWorksheetLayoutMode _worksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas;
    [ObservableProperty] private PdfHtmlProfile _htmlProfile = PdfHtmlProfile.PositionedReview;
    [ObservableProperty] private decimal _rasterDpi = 144;
    [ObservableProperty] private bool _compressPdfOutput;

    partial void OnStateChanged(ConversionJobState value) => OnPropertyChanged(nameof(CanEditOptions));

    internal OfficeWorkflowConversionOptions CreateConversionOptions() => new() {
        PageRanges = SupportsPageSelection ? PageRanges : null,
        WordMode = IsWordImport ? WordMode : null,
        PowerPointMode = IsPowerPointImport ? PowerPointMode : null,
        WorksheetLayout = IsWorksheetExport ? WorksheetLayout : null,
        HtmlProfile = IsHtmlExport ? HtmlProfile : null,
        RasterDpi = UsesRasterPages ? (double)RasterDpi : null,
        CompressPdfOutput = SupportsPdfCompression && CompressPdfOutput
    };
}
