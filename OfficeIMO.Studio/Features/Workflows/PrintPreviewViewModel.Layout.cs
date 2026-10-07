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
            new(PdfPrintAlignment.TopLeft, T("Alignment.TopLeft", "Top left")),
            new(PdfPrintAlignment.Top, T("Alignment.Top", "Top center")),
            new(PdfPrintAlignment.TopRight, T("Alignment.TopRight", "Top right")),
            new(PdfPrintAlignment.Left, T("Alignment.Left", "Middle left")),
            new(PdfPrintAlignment.Center, T("Alignment.Center", "Center")),
            new(PdfPrintAlignment.Right, T("Alignment.Right", "Middle right")),
            new(PdfPrintAlignment.BottomLeft, T("Alignment.BottomLeft", "Bottom left")),
            new(PdfPrintAlignment.Bottom, T("Alignment.Bottom", "Bottom center")),
            new(PdfPrintAlignment.BottomRight, T("Alignment.BottomRight", "Bottom right"))
        ];
        PageSubsetChoices = [
            new(PdfPrintPageSubset.All, T("PageSubset.All", "All selected pages")),
            new(PdfPrintPageSubset.Odd, T("PageSubset.Odd", "Odd source pages")),
            new(PdfPrintPageSubset.Even, T("PageSubset.Even", "Even source pages"))
        ];
        ColorChoices = [
            new(PdfPrintColorMode.Color, T("Color.Color", "Color")),
            new(PdfPrintColorMode.Grayscale, T("Color.Grayscale", "Grayscale"))
        ];
        SelectedAlignment = AlignmentChoices.Single(choice => choice.Value == PdfPrintAlignment.Center);
        SelectedPageSubset = PageSubsetChoices[0];
        SelectedColor = ColorChoices[0];
    }
}
