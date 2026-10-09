using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Native PageSheet distance cells paired with the loaded model's canonical snapshot.</summary>
internal sealed class VisioPageSheetLengthCells {
    private static readonly HashSet<string> Names = new(StringComparer.Ordinal) {
        "PageWidth", "PageHeight", "PageScale", "DrawingScale",
        "PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin",
        "BlockSizeX", "BlockSizeY", "AvenueSizeX", "AvenueSizeY",
        "LineToLineX", "LineToLineY", "LineToNodeX", "LineToNodeY"
    };

    internal VisioPageSheetLengthCells(XElement source, XElement baseline) {
        Source = Extract(source, requireValidCache: true);
        Baseline = Extract(baseline);
    }

    internal XElement Source { get; }
    internal XElement Baseline { get; }
    internal VisioPageSheetLengthCells Clone() => new(Source, Baseline);

    internal static XElement Extract(XElement sheet, bool requireValidCache = false) => new(sheet.Name,
        sheet.Elements(sheet.Name.Namespace + "Cell")
            .Where(cell => Names.Contains((string?)cell.Attribute("N") ?? string.Empty))
            .GroupBy(cell => (string)cell.Attribute("N")!, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .Select(group => group.Single())
            .Where(cell => !requireValidCache || HasValidCache(cell))
            .Select(cell => new XElement(cell)));

    private static bool HasValidCache(XElement cell) {
        if (!double.TryParse((string?)cell.Attribute("V"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
            double.IsNaN(value) || double.IsInfinity(value)) return false;
        string name = (string)cell.Attribute("N")!;
        return name == "PageWidth" || name == "PageHeight" || name == "PageScale" || name == "DrawingScale"
            ? value > 0 : value >= 0;
    }
}
