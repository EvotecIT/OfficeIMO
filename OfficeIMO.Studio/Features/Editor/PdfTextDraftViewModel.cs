using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Editor;

public sealed record TextFontChoice(PdfStandardFont? Value, string Label);
public sealed record TextFitChoice(PdfTextRegionWidthPolicy Value, string Label);

/// <summary>Editable text and fidelity choices shared by the on-page editor and its inspector.</summary>
public sealed partial class PdfTextDraftViewModel : ObservableObject {
    internal PdfTextDraftViewModel(string text, ICommand preview, ICommand cancel, IStudioLocalizer localizer) {
        _text = text;
        PreviewCommand = preview;
        CancelCommand = cancel;
        Fonts = new[] { new TextFontChoice(null, localizer.GetOrDefault("TextEdit.FontDetected", "Closest detected font")) }
            .Concat(Enum.GetValues<PdfStandardFont>().Select(font => new TextFontChoice(font, font.ToString()))).ToArray();
        Fits = [
            new(PdfTextRegionWidthPolicy.RejectOverflow, localizer.GetOrDefault("TextEdit.FitReject", "Reject overflow")),
            new(PdfTextRegionWidthPolicy.ShrinkToFit, localizer.GetOrDefault("TextEdit.FitShrink", "Shrink to fit")),
            new(PdfTextRegionWidthPolicy.PreserveFontSize, localizer.GetOrDefault("TextEdit.FitFlow", "Keep size and reflow nearby text"))
        ];
        _font = Fonts[0];
        _fit = Fits[0];
    }

    public IReadOnlyList<TextFontChoice> Fonts { get; }
    public IReadOnlyList<TextFitChoice> Fits { get; }
    public ICommand PreviewCommand { get; }
    public ICommand CancelCommand { get; }
    [ObservableProperty] private string _text;
    [ObservableProperty] private TextFontChoice _font;
    [ObservableProperty] private TextFitChoice _fit;
    [ObservableProperty] private double? _fontSize;
    [ObservableProperty] private double _minimumFontSize = 6;
    [ObservableProperty] private string _sourceStyle = string.Empty;
    [ObservableProperty] private bool _isReady;

    internal PdfTextEditOptions CaptureOptions() => new() {
        Font = Font.Value, FontSize = FontSize, RegionWidthPolicy = Fit.Value, MinimumFontSize = MinimumFontSize
    };
}

/// <summary>One reviewable occurrence from the current document search.</summary>
public sealed partial class PdfTextReplacementMatchViewModel : ObservableObject {
    internal PdfTextReplacementMatchViewModel(int index, PdfTextMatch match) { Index = index; Match = match; }
    internal int Index { get; }
    internal PdfTextMatch Match { get; }
    public int PageNumber => Match.PageNumber;
    public string Text => Match.Text;
    [ObservableProperty] private bool _isIncluded = true;
}
