using OfficeIMO.Visio;

namespace OfficeIMO.Scaling.Benchmarks;

internal sealed class VisioBuildLane : ScalingLane {
    public override string Name => "Visio.Build";
    public override string Unit => "shape+connector";

    public override object? Run() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Page-1");
        VisioShape? previous = null;
        for (int i = 0; i < Units; i++) {
            VisioShape shape = page.AddRectangle(
                OfficeLength.Inches(1 + i % 10),
                OfficeLength.Inches(1 + i / 10),
                OfficeLength.Inches(1),
                OfficeLength.Inches(0.5),
                $"Shape {i}");
            if (previous != null) {
                page.AddConnector(previous, shape);
            }

            previous = shape;
        }

        return document.ToBytes();
    }

    public override void Validate(object? result) {
        VisioDocument loaded = VisioDocument.Load(new MemoryStream((byte[])result!));
        Require(loaded.Pages[0].Shapes.Count == Units, "shape count");
        Require(loaded.Pages[0].Connectors.Count == Units - 1, "connector count");
    }

    internal static byte[] Build(int units) {
        var lane = new VisioBuildLane();
        lane.Setup(units);
        return (byte[])lane.Run()!;
    }
}

internal sealed class VisioReadLane : ScalingLane {
    private byte[] _bytes = [];
    public override string Name => "Visio.Read";
    public override string Unit => "shape+connector";
    protected override void Prepare(int units) => _bytes = VisioBuildLane.Build(units);

    public override object? Run() {
        VisioDocument loaded = VisioDocument.Load(new MemoryStream(_bytes, writable: false));
        return loaded.Pages[0].Shapes.Count;
    }

    public override void Validate(object? result) => Require((int)result! == Units, "shape count");
}
