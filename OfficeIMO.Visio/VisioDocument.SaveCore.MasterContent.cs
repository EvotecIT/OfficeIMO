using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    // A native blueprint remains native even after its last condition/error marker is edited
    // away. Use preserved content and graph ownership to choose the complete writer.
    internal static bool RequiresCompleteMasterShape(VisioShape shape) => shape.Children.Count > 0 ||
        shape.PreservedShapeChildren.Count > 0 || shape.PreservedCellElements.Count > 0 ||
        shape.PreservedNonGeometrySections.Count > 0 || shape.PreservedGeometrySections.Count > 0 ||
        shape.PreservedDataRows.Count > 0 || shape.CharacterSectionSource != null || shape.ParagraphSectionSource != null;

    /// <summary>
    /// Determines whether an instance needs an independent copy of its master's modeled content.
    /// A simple styled master still uses its specialized geometry writer, but its instances need
    /// cloned styles so their cached text frames scale without changing the blueprint.
    /// </summary>
    internal static bool RequiresMasterInstanceCopy(VisioMaster master) => master.RawMasterContentXml != null ||
        master.Shape.TextStyle != null || RequiresCompleteMasterShape(master.Shape);

    /// <summary>
    /// Creates independent effective frame caches for the generated simple 2D master path.
    /// Native, complete and dynamic connector masters retain their own frame contracts.
    /// </summary>
    internal static VisioTextStyle? CreateSimpleMasterTextFrameStyle(VisioMaster master, double width, double height) {
        if (master.RawMasterContentXml != null || RequiresCompleteMasterShape(master.Shape) ||
            (master.LoadedModelShapeXml == null && (master.ForeignResources.Count > 0 || ShapeForeignResources(new[] { master.Shape }).Any())))
            return null;
        if (TryGetBuiltinMasterDefinition(master.NameU, out var definition) && definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector)
            return null;
        return ResolveSimpleMasterTextFrameStyle(width, height, master.Shape.TextStyle);
    }

    private static VisioTextStyle ResolveSimpleMasterTextFrameStyle(double width, double height, VisioTextStyle? textStyle) {
        width = width > 0 ? width : 1;
        height = height > 0 ? height : 1;
        VisioTextStyle effective = textStyle?.Clone() ?? new VisioTextStyle();
        effective.TextPinX ??= width / 2;
        effective.TextPinY ??= height / 2;
        effective.TextWidth ??= width * 0.875;
        effective.TextHeight ??= height * 0.75;
        effective.TextLocPinX ??= effective.TextWidth / 2;
        effective.TextLocPinY ??= effective.TextHeight / 2;
        effective.TextAngle ??= 0;
        return effective;
    }

    private void WriteLoadedMasterContent(PackagePart part, VisioMaster master) {
        XDocument content = MergeRawMasterMetadata(master.RawMasterContentXml!, master);
        ApplyLoadedMasterShapeChanges(content, master);
        // Close and flush the XML before resource writing reads its surviving references.
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using StreamWriter writer = new(stream, new UTF8Encoding(false));
        writer.Write(content.Declaration + Environment.NewLine + content.ToString(SaveOptions.DisableFormatting));
    }

    private void WriteModeledMasterContent(PackagePart part, VisioMaster master, string ns, XmlWriterSettings settings) {
        using XmlWriter writer = XmlWriter.Create(part.GetStream(FileMode.Create, FileAccess.Write), settings);
        writer.WriteStartDocument(); writer.WriteStartElement("MasterContents", ns);
        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        WritePreservedAttributes(writer, master.PreservedMasterContentAttributes);
        writer.WriteStartElement("Shapes", ns); WritePreservedAttributes(writer, master.PreservedShapesAttributes);
        var ids = BuildPersistedIdMap(new[] { master.Shape }, Array.Empty<VisioConnector>(), new Dictionary<string, VisioMaster>(),
            master.PreservedAdditionalShapeElements.SelectMany(element => element.DescendantsAndSelf(XName.Get("Shape", ns))).Attributes("ID").Select(attribute => attribute.Value));
        WriteShapeElement(writer, ns, master.Shape, ids, new Dictionary<string, VisioMaster>(), Array.Empty<PackageMasterEntry>(), new Dictionary<string, int>());
        WritePreservedElements(writer, master.PreservedAdditionalShapeElements);
        writer.WriteEndElement(); WritePreservedElements(writer, master.PreservedMasterContentElements);
        writer.WriteEndElement(); writer.WriteEndDocument();
    }

    // Simple authored masters use their registered identity to select generated geometry
    // and dynamic connector routing. Modeled styling must not change that content path.
    private void WriteSimpleMasterContent(PackagePart part, VisioMaster master, string ns, XmlWriterSettings settings) {
        using XmlWriter writer = XmlWriter.Create(part.GetStream(FileMode.Create, FileAccess.Write), settings);
        writer.WriteStartDocument();
        writer.WriteStartElement("MasterContents", ns);
        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        WritePreservedAttributes(writer, master.PreservedMasterContentAttributes);
        writer.WriteStartElement("Shapes", ns);
        WritePreservedAttributes(writer, master.PreservedShapesAttributes);
        VisioShape shape = master.Shape;
        double width = shape.Width > 0 ? shape.Width : 1;
        double height = shape.Height > 0 ? shape.Height : 1;
        double localPinX = ResolveLocalPin(shape.LocPinX, width, shape.HasExplicitLocPinX);
        double localPinY = ResolveLocalPin(shape.LocPinY, height, shape.HasExplicitLocPinY);
        TryGetBuiltinMasterDefinition(master.NameU, out var definition);
        writer.WriteStartElement("Shape", ns);
        writer.WriteAttributeString("ID", "1");
        writer.WriteAttributeString("Name", shape.Name ?? shape.NameU ?? "MasterShape");
        writer.WriteAttributeString("NameU", master.NameU);
        writer.WriteAttributeString("Type", "Shape");
        if (definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector) {
            writer.WriteAttributeString("LineStyle", "0");
            writer.WriteAttributeString("FillStyle", "0");
            writer.WriteAttributeString("TextStyle", "0");
            WriteXForm1D(writer, ns, 0, 0, 1, 0);
            WriteCell(writer, ns, "OneD", 1);
            WriteCell(writer, ns, "ObjType", 2);
            WriteCell(writer, ns, "LineWeight", shape.LineWeight);
            WriteCell(writer, ns, "LinePattern", shape.LinePattern);
            WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
            WriteCell(writer, ns, "FillPattern", 0);
            WriteCellValue(writer, ns, "FillForegnd", Color.Transparent.ToVisioHex());
            WriteCell(writer, ns, "LockHeight", 1);
            WriteCell(writer, ns, "LockCalcWH", 1);
            WriteCell(writer, ns, "GlueType", 2);
            WriteCell(writer, ns, "NoAlignBox", 1);
            WriteCell(writer, ns, "DynFeedback", 2);
            WriteCell(writer, ns, "ShapeSplittable", 1);
            WriteCell(writer, ns, "LayerMember", 0);
            WriteConnectorControlSection(writer, ns, height);
        } else {
            writer.WriteAttributeString("LineStyle", "1");
            writer.WriteAttributeString("FillStyle", "1");
            writer.WriteAttributeString("TextStyle", "1");
            WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, localPinX, localPinY, shape.Angle);
            WriteCell(writer, ns, "ObjType", 1);
            if (definition?.LockAspect == true) {
                WriteCell(writer, ns, "LockAspect", 1);
            }
            WriteCell(writer, ns, "LineWeight", shape.LineWeight);
            WriteCell(writer, ns, "LinePattern", shape.LinePattern);
            WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
            WriteCell(writer, ns, "FillPattern", shape.FillPattern);
            WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
            WriteShapeGeometry(writer, ns, shape, master.NameU, width, height);
            WriteDefaultTextBlock(writer, ns, width, height, shape.TextStyle);
        }
        WriteTextBlockCells(writer, ns, shape.TextStyle);
        WriteCell(writer, ns, "ShapeSplit", 1);
        WriteCell(writer, ns, "QuickStyleType", 2);
        WriteConnectionSection(writer, ns, shape.ConnectionPoints);
        WriteMasterUserSection(writer, ns);
        WriteMasterCharacterSection(writer, ns, shape.TextStyle);
        WriteParaSection(writer, ns, shape.TextStyle);
        WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
        WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, shapeDataRows: shape.ShapeData, sectionName: shape.ShapeDataSectionName);
        WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
        writer.WriteEndElement();
        WritePreservedElements(writer, master.PreservedAdditionalShapeElements);
        writer.WriteEndElement();
        WritePreservedElements(writer, master.PreservedMasterContentElements);
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

}
