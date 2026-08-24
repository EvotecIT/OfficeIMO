namespace OfficeIMO.Visio.Benchmarks;

internal sealed record VisioBenchmarkScale(string Name, int PageCount, int ShapesPerPage) {
    internal int ShapeCount => PageCount * ShapesPerPage;
    internal int ConnectorCount => PageCount * (ShapesPerPage - 1);
}

internal sealed record VisioBenchmarkFixture(
    VisioBenchmarkScale Scale,
    byte[] PackageBytes);

internal static class VisioBenchmarkCorpus {
    internal static readonly VisioBenchmarkScale[] Scales = [
        new("Small", 1, 25),
        new("Normal", 4, 100),
        new("Large", 8, 250)
    ];

    internal static VisioBenchmarkFixture CreateFixture(VisioBenchmarkScale scale) =>
        new(scale, CreateAndSave(scale));

    internal static byte[] CreateAndSave(VisioBenchmarkScale scale) {
        VisioDocument document = CreateDocument(scale);
        return document.ToBytes();
    }

    internal static VisioDocument CreateDocument(VisioBenchmarkScale scale) {
        VisioDocument document = VisioDocument.Create();
        document.Title = $"Benchmark {scale.Name}";
        document.Author = "OfficeIMO";
        for (int pageIndex = 0; pageIndex < scale.PageCount; pageIndex++) {
            VisioPage page = document.AddPage($"Page-{pageIndex + 1}", 20, 20);
            var shapes = new List<VisioShape>(scale.ShapesPerPage);
            for (int shapeIndex = 0; shapeIndex < scale.ShapesPerPage; shapeIndex++) {
                int column = shapeIndex % 20;
                int row = shapeIndex / 20;
                VisioShape shape = page.AddRectangle(
                    0.75 + column * 0.9,
                    19.25 - row * 0.9,
                    0.7,
                    0.45,
                    $"P{pageIndex + 1}-S{shapeIndex + 1}");
                shape.SetShapeData("Index", shapeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                shapes.Add(shape);
                if (shapeIndex != 0) {
                    page.AddConnector(shapes[shapeIndex - 1], shape, ConnectorKind.Dynamic);
                }
            }
        }
        return document;
    }
}
