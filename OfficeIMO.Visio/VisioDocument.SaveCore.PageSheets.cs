using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private void WritePageSheet(XmlWriter writer, string ns, VisioPage page) {
        if (page.PageSheetLengthCells == null) {
            WritePageSheetCore(writer, ns, page);
            return;
        }

        XElement current = CreatePageSheetModel(ns, page);
        var source = new XElement(page.PageSheetLengthCells.Source);
        MergeModeledContentChanges(source, page.PageSheetLengthCells.Baseline, VisioPageSheetLengthCells.Extract(current));
        foreach (XElement cell in source.Elements()) {
            XElement? emitted = current.Elements(cell.Name).SingleOrDefault(element =>
                (string?)element.Attribute("N") == (string?)cell.Attribute("N"));
            emitted?.ReplaceWith(new XElement(cell));
        }
        current.WriteTo(writer);
    }

    private XElement CreatePageSheetModel(string ns, VisioPage page) {
        var xml = new XDocument();
        using (XmlWriter writer = xml.CreateWriter()) WritePageSheetCore(writer, ns, page);
        return xml.Root!;
    }

    private void WritePageSheetCore(XmlWriter writer, string ns, VisioPage page) {
        writer.WriteStartElement("PageSheet", ns);
        writer.WriteAttributeString("LineStyle", "0");
        writer.WriteAttributeString("FillStyle", "0");
        writer.WriteAttributeString("TextStyle", "0");

        bool useUnits = page.DefaultUnit != VisioMeasurementUnit.Inches ||
                        page.Width != 8.26771653543307 ||
                        page.Height != 11.69291338582677;
        if (useUnits) {
            string pageUnitCode = page.DefaultUnit.ToVisioUnitCode();
            WritePageCell(writer, ns, "PageWidth", page.Width, pageUnitCode);
            WritePageCell(writer, ns, "PageHeight", page.Height, pageUnitCode);
            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724, "MM");
            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724, "MM");
        } else {
            WritePageCell(writer, ns, "PageWidth", page.Width);
            WritePageCell(writer, ns, "PageHeight", page.Height);
            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724);
            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724);
        }
        VisioScaleSetting pageScale = page.GetEffectivePageScale();
        WritePageCell(writer, ns, "PageScale", pageScale.ToInches(), pageScale.Unit.ToVisioUnitCode());
        VisioScaleSetting drawingScale = page.GetEffectiveDrawingScale();
        WritePageCell(writer, ns, "DrawingScale", drawingScale.ToInches(), drawingScale.Unit.ToVisioUnitCode());
        WritePageCell(writer, ns, "DrawingSizeType", (int)page.DrawingSizeType);
        WritePageCell(writer, ns, "DrawingScaleType", 0);
        WritePageCell(writer, ns, "InhibitSnap", page.Snap ? 0 : 1);
        WritePageCell(writer, ns, "PageLockReplace", page.PageLockReplace ? 1 : 0, "BOOL");
        WritePageCell(writer, ns, "PageLockDuplicate", page.PageLockDuplicate ? 1 : 0, "BOOL");
        WritePageCell(writer, ns, "UIVisibility", (int)page.UiVisibility);
        WritePageCell(writer, ns, "ShdwType", 0);
        WritePageCell(writer, ns, "ShdwObliqueAngle", 0);
        WritePageCell(writer, ns, "ShdwScaleFactor", 1);
        WritePageCell(writer, ns, "DrawingResizeType", page.AutoResizeDrawing ? 1 : 0);
        WritePageCell(writer, ns, "PageShapeSplit", page.AllowShapeSplitting ? 1 : 0);
        WritePagePlacementCells(writer, ns, page);
        WritePageLayoutGridCells(writer, ns, page);
        WritePageLayoutRoutingCells(writer, ns, page);
        WritePageRoutingSpacingCells(writer, ns, page);
        WritePreservedElements(writer, page.PreservedPageSheetCells);
        // For non-default page sizes, include theme/margin metadata like the asset samples
        bool hasPreservedUserSection = page.PreservedPageSheetSections.Any(section =>
            string.Equals(section.Attribute("N")?.Value, "User", StringComparison.OrdinalIgnoreCase));
        if (useUnits) {
            WritePageCell(writer, ns, "ColorSchemeIndex", 60);
            WritePageCell(writer, ns, "EffectSchemeIndex", 60);
            WritePageCell(writer, ns, "ConnectorSchemeIndex", 60);
            WritePageCell(writer, ns, "FontSchemeIndex", 60);
            WritePageCell(writer, ns, "ThemeIndex", 60);
            WriteMarginCells(writer, ns, page, useUnits);
            if (page.PrintOrientation.HasValue) {
                WritePageCell(writer, ns, "PrintPageOrientation", (int)page.PrintOrientation.Value);
            }
            if (!hasPreservedUserSection) {
                writer.WriteStartElement("Section", ns);
                writer.WriteAttributeString("N", "User");
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("N", "msvThemeOrder");
                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Value");
                writer.WriteAttributeString("V", "0");
                writer.WriteEndElement();
                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Prompt");
                writer.WriteAttributeString("V", "");
                writer.WriteAttributeString("F", "No Formula");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        } else {
            if (page.HasExplicitMargins) {
                WriteMarginCells(writer, ns, page, useUnits);
            }

            if (page.PrintOrientation.HasValue) {
                WritePageCell(writer, ns, "PrintPageOrientation", (int)page.PrintOrientation.Value);
            }
        }
        BuildLayerIndexMap(page, out List<VisioLayer> layersToWrite);
        if (layersToWrite.Count > 0) {
            WriteLayerSection(writer, ns, layersToWrite);
        }
        WritePreservedElements(writer, page.PreservedPageSheetSections);
        writer.WriteEndElement();
    }
}
