using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record PrintAlignmentChoice(PdfPrintAlignment Value, string Label) {
    public override string ToString() => Label;
}
public sealed record PrintPageSubsetChoice(PdfPrintPageSubset Value, string Label) {
    public override string ToString() => Label;
}
public sealed record PrintColorChoice(PdfPrintColorMode Value, string Label) {
    public override string ToString() => Label;
}

public sealed partial class PrintPreviewViewModel {
    [ObservableProperty] private double _customScalePercent = 100;
    [ObservableProperty] private double _marginLeft = 18;
    [ObservableProperty] private double _marginTop = 18;
    [ObservableProperty] private double _marginRight = 18;
    [ObservableProperty] private double _marginBottom = 18;
    [ObservableProperty] private PrintAlignmentChoice _selectedAlignment = null!;
    [ObservableProperty] private PrintPageSubsetChoice _selectedPageSubset = null!;
    [ObservableProperty] private PrintColorChoice _selectedColor = null!;
    public IReadOnlyList<PrintAlignmentChoice> AlignmentChoices { get; private set; } = [];
    public IReadOnlyList<PrintPageSubsetChoice> PageSubsetChoices { get; private set; } = [];
    public IReadOnlyList<PrintColorChoice> ColorChoices { get; private set; } = [];
    public bool UsesCustomScale => SelectedScale?.Value == PdfPrintScaleMode.Custom;

    private void InitializeLayoutChoices() {
        AlignmentChoices = [
            new(PdfPrintAlignment.TopLeft, T("Alignment.TopLeft")),
            new(PdfPrintAlignment.Top, T("Alignment.Top")),
            new(PdfPrintAlignment.TopRight, T("Alignment.TopRight")),
            new(PdfPrintAlignment.Left, T("Alignment.Left")),
            new(PdfPrintAlignment.Center, T("Alignment.Center")),
            new(PdfPrintAlignment.Right, T("Alignment.Right")),
            new(PdfPrintAlignment.BottomLeft, T("Alignment.BottomLeft")),
            new(PdfPrintAlignment.Bottom, T("Alignment.Bottom")),
            new(PdfPrintAlignment.BottomRight, T("Alignment.BottomRight"))
        ];
        PageSubsetChoices = [
            new(PdfPrintPageSubset.All, T("PageSubset.All")),
            new(PdfPrintPageSubset.Odd, T("PageSubset.Odd")),
            new(PdfPrintPageSubset.Even, T("PageSubset.Even"))
        ];
        ColorChoices = [
            new(PdfPrintColorMode.Color, T("Color.Color")),
            new(PdfPrintColorMode.Grayscale, T("Color.Grayscale"))
        ];
        SelectedAlignment = AlignmentChoices.Single(choice => choice.Value == PdfPrintAlignment.Center);
        SelectedPageSubset = PageSubsetChoices[0];
        SelectedColor = ColorChoices[0];
    }
}
