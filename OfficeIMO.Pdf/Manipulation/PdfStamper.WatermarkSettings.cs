using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static PdfDictionary BuildWatermarkSettings(PdfWatermarkOptions options,
        Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber) {
        var result = new PdfDictionary();
        result.Items["Version"] = new PdfNumber(1);
        result.Items["Text"] = new PdfStringObj(options.Text, useTextStringEncoding: true);
        result.Items["Width"] = new PdfNumber(options.Width);
        result.Items["Height"] = new PdfNumber(options.Height);
        result.Items["FontSize"] = new PdfNumber(options.FontSize);
        result.Items["Font"] = new PdfNumber((int)options.Font);
        result.Items["Rotation"] = new PdfNumber(options.RotationDegrees);
        result.Items["Opacity"] = new PdfNumber(options.Opacity);
        result.Items["R"] = new PdfNumber(options.Color.R);
        result.Items["G"] = new PdfNumber(options.Color.G);
        result.Items["B"] = new PdfNumber(options.Color.B);
        if (options.X is { } x) result.Items["X"] = new PdfNumber(x);
        if (options.Y is { } y) result.Items["Y"] = new PdfNumber(y);
        if (options.ImageBytes is { } image) {
            int imageNumber = nextObjectNumber++;
            objects[imageNumber] = new PdfIndirectObject(imageNumber, 0, new PdfStream(new PdfDictionary(), image));
            result.Items["Image"] = new PdfReference(imageNumber, 0);
        }
        return result;
    }

    internal static IReadOnlyList<PdfWatermarkOptions> ReadWatermarks(byte[] pdf, PdfLoadOptions? readOptions) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, readOptions);
        var document = PdfReadDocument.Open(pdf, readOptions);
        var found = new Dictionary<string, (PdfWatermarkOptions Settings, List<int> Pages)>(StringComparer.Ordinal);
        for (int index = 0; index < document.Pages.Count; index++) {
            var page = (PdfDictionary)objects[document.Pages[index].ObjectNumber].Value;
            foreach (var stream in GetPageContentStreams(objects, page)) {
                string? id = stream.Dictionary.Get<PdfStringObj>("OfficeIMOWatermarkId")?.Value;
                if (id is null) continue;
                if (!stream.Dictionary.Items.TryGetValue("OfficeIMOWatermarkSettings", out var settingsObject)
                    || PdfObjectLookup.Resolve(objects, settingsObject) is not PdfDictionary settings
                    || settings.Get<PdfNumber>("Version")?.Value != 1) continue;
                if (!found.TryGetValue(id, out var item)) {
                    double Number(string key) => settings.Get<PdfNumber>(key)?.Value
                        ?? throw new InvalidDataException("Watermark settings are incomplete.");
                    var options = new PdfWatermarkOptions {
                        Id = id, Text = settings.Get<PdfStringObj>("Text")?.Value ?? string.Empty,
                        Width = Number("Width"), Height = Number("Height"), FontSize = Number("FontSize"),
                        Font = (PdfStandardFont)(int)Number("Font"), RotationDegrees = Number("Rotation"),
                        Opacity = Number("Opacity"), Color = new PdfColor(Number("R"), Number("G"), Number("B")),
                        X = settings.Get<PdfNumber>("X")?.Value, Y = settings.Get<PdfNumber>("Y")?.Value,
                        BehindContent = stream.Dictionary.Get<PdfBoolean>("OfficeIMOWatermarkBehind")?.Value ?? false
                    };
                    if (settings.Items.TryGetValue("Image", out var imageObject)) {
                        if (PdfObjectLookup.Resolve(objects, imageObject) is not PdfStream image || image.DecodingFailed)
                            throw new InvalidDataException("Watermark image settings cannot be read.");
                        options.ImageBytes = (byte[])image.Data.Clone();
                    }
                    options.Validate();
                    found.Add(id, item = (options, new List<int>()));
                }
                if (!item.Pages.Contains(index + 1)) item.Pages.Add(index + 1);
            }
        }
        foreach (var item in found.Values)
            item.Settings.TargetPages = PdfPageSelector.Parse(string.Join(",", item.Pages.Select(page => page.ToString(CultureInfo.InvariantCulture))));
        return found.Values.Select(item => item.Settings).ToArray();
    }
}
