namespace OfficeIMO.Visio.Benchmarks;

internal static class VisioBenchmarkValidation {
    internal static void ValidateAll(bool writeSummary) {
        foreach (VisioBenchmarkScale scale in VisioBenchmarkCorpus.Scales) {
            VisioBenchmarkFixture fixture = VisioBenchmarkCorpus.CreateFixture(scale);
            ValidatePackage(fixture);
            if (writeSummary) {
                Console.WriteLine(
                    $"{scale.Name,-6} {scale.PageCount,2} pages | {scale.ShapeCount,5:N0} shapes | " +
                    $"{scale.ConnectorCount,5:N0} connectors | {fixture.PackageBytes.Length,10:N0} bytes");
            }
        }
    }

    internal static VisioInspectionSnapshot LoadAndInspect(VisioBenchmarkFixture fixture) {
        using var stream = new MemoryStream(fixture.PackageBytes, writable: false);
        VisioDocument document = VisioDocument.Load(stream);
        return document.CreateInspectionSnapshot();
    }

    internal static void ValidateBytes(VisioBenchmarkScale scale, byte[] packageBytes) =>
        ValidatePackage(new VisioBenchmarkFixture(scale, packageBytes));

    internal static void ValidatePackage(VisioBenchmarkFixture fixture) {
        VisioInspectionSnapshot snapshot = LoadAndInspect(fixture);
        if (snapshot.Pages.Count != fixture.Scale.PageCount) {
            throw new InvalidOperationException("Visio page count did not round-trip.");
        }
        int shapes = snapshot.Pages.Sum(page => page.Shapes.Count);
        int connectors = snapshot.Pages.Sum(page => page.Connectors.Count);
        if (shapes != fixture.Scale.ShapeCount || connectors != fixture.Scale.ConnectorCount) {
            throw new InvalidOperationException(
                $"Visio graph was {shapes} shapes/{connectors} connectors; expected " +
                $"{fixture.Scale.ShapeCount}/{fixture.Scale.ConnectorCount}.");
        }
        string expectedLast = $"P{fixture.Scale.PageCount}-S{fixture.Scale.ShapesPerPage}";
        if (!snapshot.Pages.SelectMany(page => page.Shapes).Any(shape => shape.Text == expectedLast)) {
            throw new InvalidOperationException("Visio boundary shape text did not round-trip.");
        }
        if (snapshot.Pages.SelectMany(page => page.Shapes).Any(shape => shape.ShapeData.Count != 1)) {
            throw new InvalidOperationException("Visio Shape Data did not round-trip.");
        }
    }
}
