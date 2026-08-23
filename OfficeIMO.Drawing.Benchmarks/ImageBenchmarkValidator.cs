namespace OfficeIMO.Drawing.Benchmarks;

internal static class ImageBenchmarkValidator {
    internal static void Validate(TextWriter writer) {
        foreach (ImageBenchmarkAsset asset in ImageBenchmarkCorpus.All) {
            byte[] encoded = asset.ReadBytes();
            ImageBenchmarkCorpus.AssertIdentified(encoded, asset.Format, asset.Width, asset.Height, asset.Id);
            if (ReferenceEquals(asset, ImageBenchmarkCorpus.Bitmap)) {
                writer.WriteLine($"{asset.Id,-14} {asset.Format,-5} {asset.Width,4}x{asset.Height,-4} {encoded.Length,10:N0} bytes metadata-only (trailing payload)");
                continue;
            }
            OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(encoded, asset.Id);
            if (decoded.Width != asset.Width || decoded.Height != asset.Height) {
                throw new InvalidOperationException($"{asset.Id} decoded to {decoded.Width}x{decoded.Height}.");
            }
            writer.WriteLine($"{asset.Id,-14} {asset.Format,-5} {asset.Width,4}x{asset.Height,-4} {encoded.Length,10:N0} bytes {ImageBenchmarkCorpus.PixelHash(decoded)[..16]}");
        }

        byte[] bmp = ImageBenchmarkCorpus.CreateBmp24();
        OfficeRasterImage decodedBmp = ImageBenchmarkCorpus.Decode(bmp, "GeneratedBmp24");
        writer.WriteLine($"{"GeneratedBmp24",-14} {OfficeImageFormat.Bmp,-5} {decodedBmp.Width,4}x{decodedBmp.Height,-4} {bmp.Length,10:N0} bytes {ImageBenchmarkCorpus.PixelHash(decodedBmp)[..16]}");

        var encode = new ImageEncodeBenchmarks();
        encode.Setup();
        writer.WriteLine($"Encode PNG     {encode.Png().Length,10:N0} bytes");
        writer.WriteLine($"Encode JPEG    {encode.Jpeg().Length,10:N0} bytes");
        writer.WriteLine($"Encode TIFF    {encode.Tiff().Length,10:N0} bytes");
        writer.WriteLine($"Encode WebP    {encode.Webp().Length,10:N0} bytes");

        var transform = new ImageTransformBenchmarks();
        transform.Setup();
        OfficeImageOptimizationResult optimized = transform.OptimizeForPlacement();
        writer.WriteLine($"Optimize JPEG  {optimized.Final.Width,4}x{optimized.Final.Height,-4} {optimized.FinalEncodedLength,10:N0} bytes {optimized.Status}");
    }
}
