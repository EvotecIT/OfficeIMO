using System.IO.Compression;
using System.Xml.Linq;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>Reads and saves document properties from a presentation with missing property values.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointDocumentPropertiesBenchmarks {
    private const int SlideCount = 20;
    private byte[] _deck = Array.Empty<byte>();
    private byte[] _complete = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_PROPERTIES_FIXTURE");
        if (string.IsNullOrWhiteSpace(fixture)) {
            _deck = CreateFixture(out _complete);
        } else {
            _deck = File.ReadAllBytes(fixture);
            _complete = File.ReadAllBytes(Path.ChangeExtension(fixture, ".complete.pptx"));
        }
        if (ReadProperties() != ("OfficeIMO", "20")) {
            throw new InvalidOperationException("The property benchmark did not read the expected metadata.");
        }
        using var saved = PowerPointDocument.Load(new MemoryStream(SaveMissingProperties(), writable: false));
        if (saved.Slides.Count != SlideCount || saved.ApplicationProperties.TotalTime != "0" || saved.Properties.Modified == null) {
            throw new InvalidOperationException("The save did not fill in the missing properties.");
        }
        using var complete = PowerPointDocument.Load(new MemoryStream(SaveCompleteProperties(), writable: false));
        if (complete.Slides.Count != SlideCount || complete.ApplicationProperties.Slides != "20" || complete.Properties.Modified == null) {
            throw new InvalidOperationException("The save did not preserve complete properties.");
        }
        using var edited = PowerPointDocument.Load(new MemoryStream(EditAndSaveProperties(), writable: false));
        if (edited.Properties.Title != "Quarterly review" || edited.ApplicationProperties.Company != "Evotec") {
            throw new InvalidOperationException("The property benchmark did not save the expected changes.");
        }
    }

    internal static byte[] CreateFixture(out byte[] complete) {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Quarterly review");
        }

        complete = document.ToBytes();
        using var original = new ZipArchive(new MemoryStream(complete, writable: false));
        using var result = new MemoryStream();
        using (var output = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in original.Entries) {
                using Stream source = entry.Open();
                using Stream destination = output.CreateEntry(entry.FullName).Open();
                string? missing = entry.FullName switch {
                    "docProps/core.xml" => "modified",
                    "docProps/app.xml" => "TotalTime",
                    _ => null
                };
                if (missing == null) {
                    source.CopyTo(destination);
                } else {
                    XDocument xml = XDocument.Load(source);
                    xml.Root!.Elements().Where(element => element.Name.LocalName == missing).Remove();
                    xml.Save(destination);
                }
            }
        }
        return result.ToArray();
    }

    [Benchmark]
    public (string?, string) ReadProperties() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return (document.Properties.Creator, document.ApplicationProperties.Slides);
    }

    [Benchmark]
    public byte[] SaveMissingProperties() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] SaveCompleteProperties() {
        using var document = PowerPointDocument.Load(new MemoryStream(_complete, writable: false));
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] EditAndSaveProperties() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        document.Properties.Title = "Quarterly review";
        document.ApplicationProperties.Company = "Evotec";
        return document.ToBytes();
    }
}
