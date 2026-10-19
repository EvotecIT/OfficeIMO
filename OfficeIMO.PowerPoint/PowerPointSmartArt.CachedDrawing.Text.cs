using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSmartArt {
    /// <summary>Stages a text-only update to a current producer cache, preserving its shape geometry and text style.</summary>
    private bool TryPrepareCachedNodeTextEdit(XDocument data, XElement sourceBody, string text,
        out DiagramPersistLayoutPart? part, out XDocument? drawing) {
        part = FindCachedDrawingPart(data);
        drawing = null;
        string? sourceId = (string?)sourceBody.Parent?.Attribute("modelId");
        OpenXmlCompositeElement? tree = part?.Drawing?.ChildElements.OfType<OpenXmlCompositeElement>()
            .FirstOrDefault(element => element.LocalName == "spTree");
        if (tree == null || string.IsNullOrEmpty(sourceId) ||
            !MatchesDiagramText(data, tree, out Dictionary<string, string> associations)) return false;

        XNamespace dsp = PersistedDiagramNamespace;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XDocument candidate = XDocument.Parse(part!.Drawing!.OuterXml, LoadOptions.PreserveWhitespace);
        int updated = 0;
        foreach (XElement shape in candidate.Descendants(dsp + "sp")) {
            string? id = (string?)shape.Attribute("modelId");
            if (id != null && associations.TryGetValue(id, out string? semantic)) id = semantic;
            if (!string.Equals(id, sourceId, StringComparison.OrdinalIgnoreCase)) continue;
            XElement? body = shape.Element(dsp + "txBody");
            XElement? paragraph = body?.Element(a + "p");
            if (body == null || paragraph == null) return false;
            // SetNodeText replaces the node's paragraph content. Retain the producer's
            // paragraph and first-run style when projecting that plain text into its cache.
            var replacement = new XElement(a + "p");
            if (paragraph.Element(a + "pPr") is XElement properties) replacement.Add(new XElement(properties));
            var run = new XElement(a + "r");
            if (paragraph.Elements(a + "r").FirstOrDefault()?.Element(a + "rPr") is XElement runStyle)
                run.Add(new XElement(runStyle));
            run.Add(new XElement(a + "t", text));
            replacement.Add(run);
            if (paragraph.Element(a + "endParaRPr") is XElement endStyle) replacement.Add(new XElement(endStyle));
            body.Elements(a + "p").Remove();
            body.Add(replacement);
            updated++;
        }
        if (updated == 0) return false;
        drawing = candidate;
        return true;
    }
}
