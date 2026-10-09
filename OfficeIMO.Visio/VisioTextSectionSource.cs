using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Native text-row syntax paired with the loaded model's canonical value snapshot.</summary>
internal sealed class VisioTextSectionSource {
    internal VisioTextSectionSource(XElement source, XElement baseline, bool inherited = false) {
        Source = new XElement(source);
        Baseline = new XElement(baseline);
        Inherited = inherited;
    }

    internal XElement Source { get; }
    internal XElement Baseline { get; }
    internal bool Inherited { get; }
    internal VisioTextSectionSource Clone() => new(Source, Baseline, Inherited);
}
