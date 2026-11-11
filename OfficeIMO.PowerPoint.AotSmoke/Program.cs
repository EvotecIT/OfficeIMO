using OfficeIMO;
using System.Collections;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

HiddenDataSmoke.Run();

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".pptx");
string mastersPath = Path.ChangeExtension(path, ".masters.pptx");
try {
    using (PowerPointDocument presentation = PowerPointDocument.Create(mastersPath)) {
        PowerPointSlideMaster master = presentation.SlideMasters[0];
        master.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0),
            OfficeLength.Points(720), OfficeLength.Points(12), "Band").Fill(OfficeColor.Parse("1F4E79"));
        PowerPointSlideLayout layout = master.AddLayout("AOT layout");
        layout.AddPlaceholder(PowerPointPlaceholderType.Title, new PowerPointLayoutBox(
            OfficeLength.Points(40), OfficeLength.Points(30), OfficeLength.Points(600), OfficeLength.Points(50)));
        PowerPointSlideLayout copy = layout.Duplicate();
        PowerPointSlide slide = presentation.AddSlide(copy);
        slide.AddTextBox("Native AOT").BindToPlaceholder(PowerPointPlaceholderType.Title);
        PowerPointAutoShape from = slide.AddRectangle(OfficeLength.Points(40), OfficeLength.Points(150), OfficeLength.Points(100), OfficeLength.Points(50));
        PowerPointAutoShape to = slide.AddShape(OfficePresetShapeType.Ellipse, OfficeLength.Points(400), OfficeLength.Points(300),
            OfficeLength.Points(100), OfficeLength.Points(60));
        PowerPointConnectionShape connector = slide.AddConnector(from, to, PowerPointConnectorKind.Elbow);
        to.MoveBy(OfficeLength.Points(20), OfficeLength.Points(0));
        copy.Remove(layout);
        if (connector.EndShape != to || slide.Layout != layout || presentation.ValidateDocument().Count != 0) {
            throw new InvalidOperationException("PowerPoint master, layout, or connector authoring failed under NativeAOT.");
        }
        presentation.Save();
    }

    using (PowerPointDocument reloaded = PowerPointDocument.Load(mastersPath)) {
        PowerPointConnectionShape connector = reloaded.Slides[0].Shapes.OfType<PowerPointConnectionShape>().Single();
        if (connector.Kind != PowerPointConnectorKind.Elbow || connector.StartShape == null || connector.EndShape == null ||
            !reloaded.SlideMasters[0].Shapes.Any(shape => shape.Name == "Band")) {
            throw new InvalidOperationException("The PowerPoint master or connector glue did not survive a NativeAOT round trip.");
        }
    }

    Console.WriteLine("PASS | PowerPoint master shapes, custom layouts, placeholder binding, and glued connectors");

    using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeSeries("Revenue", new[] { 1d, 2d, 3d }) });

        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("OfficeIMO NativeAOT slide");
        slide.AddChart(OfficeChartType.ColumnClustered, data);
        var dictionaryRows = new[] {
            new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1)),
            new GenericOnlyDictionaryRow<int>(("Score", 20), ("Rank", 2))
        };
        PowerPointTable dictionaryTable = slide.AddTable(dictionaryRows);
        if (dictionaryTable.GetCell(0, 0).Text != "Score" || dictionaryTable.GetCell(0, 1).Text != "Rank" ||
            dictionaryTable.GetCell(1, 0).Text != "10" || dictionaryTable.GetCell(1, 1).Text != "1") {
            throw new InvalidOperationException("PowerPoint AddTable did not preserve generic-only dictionary columns under NativeAOT.");
        }
        presentation.DuplicateSlide(0);
        presentation.Save();
    }

    using PowerPointDocument reopened = PowerPointDocument.Load(path);
    if (reopened.Slides.Count != 2 || reopened.Slides[0].Charts.Count() != 1 || reopened.Slides[1].Charts.Count() != 1 ||
        reopened.Slides[0].Tables.Single().GetCell(1, 0).Text != "10" || reopened.Slides[1].Tables.Single().GetCell(1, 0).Text != "10") {
        throw new InvalidOperationException("The PowerPoint round trip lost its slide or cloned chart relationships.");
    }

    Console.WriteLine("PASS | PowerPoint chart and generic-only dictionary table create, duplicate, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
    if (File.Exists(mastersPath)) File.Delete(mastersPath);
}

file sealed class GenericOnlyDictionaryRow<TValue> : IReadOnlyDictionary<string, TValue> {
    private readonly KeyValuePair<string, TValue>[] _entries;
    private readonly Dictionary<string, TValue> _lookup;

    internal GenericOnlyDictionaryRow(params (string Key, TValue Value)[] entries) {
        _entries = entries.Select(entry => new KeyValuePair<string, TValue>(entry.Key, entry.Value)).ToArray();
        _lookup = _entries.ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.OrdinalIgnoreCase);
    }

    public TValue this[string key] => _lookup[key];
    public IEnumerable<string> Keys => _entries.Select(entry => entry.Key);
    public IEnumerable<TValue> Values => _entries.Select(entry => entry.Value);
    public int Count => _entries.Length;
    public bool ContainsKey(string key) => _lookup.ContainsKey(key);
    public bool TryGetValue(string key, out TValue value) => _lookup.TryGetValue(key, out value!);
    public IEnumerator<KeyValuePair<string, TValue>> GetEnumerator() => ((IEnumerable<KeyValuePair<string, TValue>>)_entries).GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
