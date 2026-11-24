using OfficeIMO.OpenDocument;

namespace OfficeIMO.Scaling.Benchmarks;

internal sealed class OdtBuildLane : ScalingLane {
    public override string Name => "Odt.Build";
    public override string Unit => "paragraph";

    public override object? Run() {
        var document = OdtDocument.Create();
        for (int i = 0; i < Units; i++) {
            document.AddParagraph($"Paragraph {i} with some ordinary text.");
        }

        return document.ToBytes();
    }

    public override void Validate(object? result) =>
        Require(OdtDocument.Load(new MemoryStream((byte[])result!)).Paragraphs.Count == Units, "paragraph count");

    internal static byte[] Build(int units) {
        var lane = new OdtBuildLane();
        lane.Setup(units);
        return (byte[])lane.Run()!;
    }
}

internal sealed class OdtReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Odt.Read";
    public override string Unit => "paragraph";
    protected override void Prepare(int units) => _bytes = OdtBuildLane.Build(units);
    public override object? Run() => OdtDocument.Load(new MemoryStream(_bytes, writable: false)).Paragraphs.Count;
    public override void Validate(object? result) => Require((int)result! == Units, "paragraph count");
}

internal sealed class OdsBuildLane : ScalingLane {
    public override string Name => "Ods.Build";
    public override string Unit => "row";

    public override object? Run() {
        var document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int row = 0; row < Units; row++) {
            sheet.Cell(row, 0).SetNumber(row);
            for (int column = 1; column < 5; column++) {
                sheet.Cell(row, column).SetString($"r{row}c{column}");
            }
        }

        return document.ToBytes();
    }

    public override void Validate(object? result) =>
        Require(OdsDocument.Load(new MemoryStream((byte[])result!)).Sheets[0].RowCount == Units, "row count");

    internal static byte[] Build(int units) {
        var lane = new OdsBuildLane();
        lane.Setup(units);
        return (byte[])lane.Run()!;
    }
}

internal sealed class OdsReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Ods.Read";
    public override string Unit => "row";
    protected override void Prepare(int units) => _bytes = OdsBuildLane.Build(units);
    public override object? Run() => OdsDocument.Load(new MemoryStream(_bytes, writable: false)).Sheets[0].RowCount;
    public override void Validate(object? result) => Require(Convert.ToInt64(result) == Units, "row count");
}

internal sealed class OdpBuildLane : ScalingLane {
    public override string Name => "Odp.Build";
    public override string Unit => "slide";

    public override object? Run() {
        var presentation = OdpPresentation.Create();
        for (int i = 0; i < Units; i++) {
            presentation.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 2), $"Slide {i}");
        }

        return presentation.ToBytes();
    }

    public override void Validate(object? result) =>
        Require(OdpPresentation.Load(new MemoryStream((byte[])result!)).Slides.Count == Units, "slide count");

    internal static byte[] Build(int units) {
        var lane = new OdpBuildLane();
        lane.Setup(units);
        return (byte[])lane.Run()!;
    }
}

internal sealed class OdpReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Odp.Read";
    public override string Unit => "slide";
    protected override void Prepare(int units) => _bytes = OdpBuildLane.Build(units);
    public override object? Run() => OdpPresentation.Load(new MemoryStream(_bytes, writable: false)).Slides.Count;
    public override void Validate(object? result) => Require((int)result! == Units, "slide count");
}
