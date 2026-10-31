using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.PowerPoint;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Workflows;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionJobViewModel {
    public bool SupportsPageSelection => Route.Route.SupportsPageSelection;
    public bool SupportsPdfCompression => Route.Route.SupportsPdfCompression;
    public bool IsWordImport => Route.Route.Id == "pdf-docx";
    public bool IsPowerPointImport => Route.Route.Id == "pdf-pptx";
    public bool IsWorksheetExport => Route.Route.Id == "xlsx-pdf";
    public bool IsHtmlExport => Route.Route.Id == "pdf-html";
    public bool IsLegacyWordExport => Route.Route.Id == "doc-pdf";
    public bool IsPlainTextExport => Route.Route.Id == "txt-pdf";
    public bool CanEditOptions => State == ConversionJobState.Queued || CanRetry;
    public bool UsesRasterPages => IsWordImport && WordMode == WordPdfImportMode.VisualPages ||
        IsPowerPointImport && PowerPointMode is PowerPointPdfImportMode.VisualPages or PowerPointPdfImportMode.HybridVisualAndEditableTables;
    public IReadOnlyList<WordPdfImportMode> WordModes { get; } = Enum.GetValues<WordPdfImportMode>();
    public IReadOnlyList<PowerPointPdfImportMode> PowerPointModes { get; } = [
        PowerPointPdfImportMode.EditableContent, PowerPointPdfImportMode.VisualPages,
        PowerPointPdfImportMode.HybridVisualAndEditableTables, PowerPointPdfImportMode.EditableTables];
    public IReadOnlyList<ExcelPdfWorksheetLayoutMode> WorksheetLayouts { get; } = Enum.GetValues<ExcelPdfWorksheetLayoutMode>();
    public IReadOnlyList<PdfHtmlProfile> HtmlProfiles { get; } = Enum.GetValues<PdfHtmlProfile>();

    [ObservableProperty] private string _pageRanges = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(UsesRasterPages)), NotifyPropertyChangedFor(nameof(Fidelity))]
    private WordPdfImportMode _wordMode = WordPdfImportMode.EditableContent;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(UsesRasterPages)), NotifyPropertyChangedFor(nameof(Fidelity))]
    private PowerPointPdfImportMode _powerPointMode = PowerPointPdfImportMode.EditableContent;
    [ObservableProperty] private ExcelPdfWorksheetLayoutMode _worksheetLayout = ExcelPdfWorksheetLayoutMode.WorksheetCanvas;
    [ObservableProperty] private PdfHtmlProfile _htmlProfile = PdfHtmlProfile.PositionedReview;
    [ObservableProperty] private decimal _rasterDpi = 144;
    [ObservableProperty] private bool _compressPdfOutput;
    [ObservableProperty] private bool _allowLegacyImportLoss;
    [ObservableProperty] private string _textEncoding = string.Empty;
    [ObservableProperty] private decimal _tabSize = 8;

    partial void OnStateChanged(ConversionJobState value) {
        OnPropertyChanged(nameof(CanEditOptions));
        NotifyOutcome();
    }

    internal OfficeConversionOptions CreateConversionOptions() {
        var options = new OfficeConversionOptions {
            Word = Route.Route.Id is "docx-pdf" or "doc-pdf" ? new WordPdfExportOptions().UseProfile(PdfProfile) : null,
            Excel = IsWorksheetExport ? new ExcelPdfExportOptions().UseProfile(PdfProfile) : null,
            PowerPoint = Route.Route.Id == "pptx-pdf" ? new PowerPointPdfExportOptions().UseProfile(PdfProfile) : null,
            PdfLoad = SupportsPageSelection && !string.IsNullOrWhiteSpace(PageRanges)
                ? new PdfLoadOptions { PageSelection = PdfPageSelection.Parse(PageRanges) } : null,
            PdfWord = IsWordImport ? new WordPdfImportOptions { Mode = WordMode, Dpi = (double)RasterDpi } : null,
            PdfPowerPoint = IsPowerPointImport ? new PowerPointPdfImportOptions { Mode = PowerPointMode, Dpi = (double)RasterDpi } : null,
            PdfHtml = IsHtmlExport ? new PdfToHtmlOptions { Profile = HtmlProfile } : null,
            CompressPdfOutput = SupportsPdfCompression && CompressPdfOutput,
            WordLoad = IsLegacyWordExport ? new OfficeIMO.Word.WordLoadOptions { FailOnLoss = !AllowLegacyImportLoss } : null,
            PlainText = IsPlainTextExport ? new PdfTextImportOptions { EncodingName = string.IsNullOrWhiteSpace(TextEncoding) ? null : TextEncoding, TabSize = (int)TabSize } : null
        };
        if (options.Excel != null) options.Excel.WorksheetLayout = WorksheetLayout;
        return options;
    }

    private PdfExportProfile PdfProfile => OutputProfile switch {
        OfficeConversionProfile.Faithful => PdfExportProfile.Faithful,
        OfficeConversionProfile.Lightweight => PdfExportProfile.Lightweight,
        OfficeConversionProfile.PrintReady => PdfExportProfile.PrintReady,
        OfficeConversionProfile.TextOnly => PdfExportProfile.TextOnly,
        _ => throw new ArgumentOutOfRangeException(nameof(OutputProfile), OutputProfile, "Unsupported PDF export profile.")
    };

}
