using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private void WriteShapeElement(XmlWriter writer, string ns, VisioShape shape, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyDictionary<string, VisioMaster> effectiveMasters, IReadOnlyList<PackageMasterEntry> packageMasters, IReadOnlyDictionary<string, int> layerIndexes) {
            writer.WriteStartElement("Shape", ns);
            writer.WriteAttributeString("ID", GetPersistedId(persistedIds, shape.Id));
            string shapeName = shape.Name ?? shape.NameU ?? $"Shape{shape.Id}";
            writer.WriteAttributeString("Name", shapeName);
            VisioMaster? effectiveMaster = TryGetEffectiveMaster(effectiveMasters, shape);
            writer.WriteAttributeString("NameU", shape.NameU ?? effectiveMaster?.NameU ?? shapeName);

            bool isRawMasterBackedShape = effectiveMaster?.RawMasterContentXml != null;
            bool useLocalGeometryForGeneratedStencil = effectiveMaster != null &&
                                                       effectiveMaster.RawMasterContentXml == null &&
                                                       VisioStencilMetadata.HasStencilMetadata(shape);
            bool isRawMasterGroup = isRawMasterBackedShape &&
                                    string.Equals(effectiveMaster!.Shape.Type, "Group", StringComparison.OrdinalIgnoreCase);
            bool isGroup = string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase) ||
                           shape.Children.Count > 0 ||
                           isRawMasterGroup;
            writer.WriteAttributeString("Type", isGroup ? "Group" : "Shape");
            if (!isRawMasterBackedShape) {
                writer.WriteAttributeString("LineStyle", "0");
                writer.WriteAttributeString("FillStyle", "0");
                writer.WriteAttributeString("TextStyle", "0");
            }

            if (effectiveMaster != null) {
                writer.WriteAttributeString("Master", GetPackageMasterId(packageMasters, effectiveMaster));
                if (!useLocalGeometryForGeneratedStencil && shape.MasterShapeId != null) {
                    writer.WriteAttributeString("MasterShape", shape.MasterShapeId);
                }
            }

            KeyValuePair<string, string>? originalIdEntry = GetOriginalIdEntry(persistedIds, shape.Id);
            bool wroteChildShapesInBody = false;
            if (effectiveMaster != null && !useLocalGeometryForGeneratedStencil && (isRawMasterBackedShape || !isGroup)) {
                WriteMasterBackedShapeBody(writer, ns, shape, effectiveMaster, originalIdEntry, persistedIds, layerIndexes);
                wroteChildShapesInBody = isRawMasterBackedShape;
            } else if (shape.PreservedShapeChildren.Count > 0) {
                wroteChildShapesInBody = WriteStandaloneShapeBodyWithPreservedChildOrder(writer, ns, shape, isGroup, originalIdEntry, persistedIds, effectiveMasters, packageMasters, layerIndexes);
            } else {
                WriteStandaloneShapeBody(writer, ns, shape, isGroup, originalIdEntry, persistedIds, layerIndexes);
            }

            if (!wroteChildShapesInBody && isGroup && shape.Children.Count > 0) {
                WriteChildShapesContainer(writer, ns, shape, persistedIds, effectiveMasters, packageMasters, layerIndexes);
            }

            writer.WriteEndElement();
        }

        private void WriteMasterBackedShapeBody(XmlWriter writer, string ns, VisioShape shape, VisioMaster master, KeyValuePair<string, string>? originalIdEntry, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyDictionary<string, int> layerIndexes) {
            if (WriteMasterDeltasOnly && master.RawMasterContentXml != null) {
                double masterWidth = master.Shape.Width > 0 ? master.Shape.Width : 1;
                double masterHeight = master.Shape.Height > 0 ? master.Shape.Height : 1;
                double width = shape.Width > 0 ? shape.Width : masterWidth;
                double height = shape.Height > 0 ? shape.Height : masterHeight;
                double locPinX = Math.Abs(shape.LocPinX) < double.Epsilon ? width / 2 : shape.LocPinX;
                double locPinY = Math.Abs(shape.LocPinY) < double.Epsilon ? height / 2 : shape.LocPinY;
                double masterLocPinX = Math.Abs(master.Shape.LocPinX) < double.Epsilon ? masterWidth / 2 : master.Shape.LocPinX;
                double masterLocPinY = Math.Abs(master.Shape.LocPinY) < double.Epsilon ? masterHeight / 2 : master.Shape.LocPinY;
                bool sizeDiffers = Math.Abs(width - masterWidth) > 1e-12 ||
                                   Math.Abs(height - masterHeight) > 1e-12;
                bool locPinDiffers = Math.Abs(locPinX - masterLocPinX) > 1e-12 ||
                                     Math.Abs(locPinY - masterLocPinY) > 1e-12;
                if (sizeDiffers || locPinDiffers || Math.Abs(shape.Angle) > 1e-12) {
                    WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, locPinX, locPinY, shape.Angle);
                } else {
                    WriteCell(writer, ns, "PinX", shape.PinX);
                    WriteCell(writer, ns, "PinY", shape.PinY);
                }

                WriteCell(writer, ns, "ObjType", 1);
                WriteLayerMemberCell(writer, ns, shape.LayerNames, layerIndexes);
                WriteRelationshipCell(writer, ns, shape, persistedIds);
                WriteShapeLayoutCells(writer, ns, shape);
                WriteProtectionCells(writer, ns, shape.Protection);
                WritePreservedElements(writer, shape.PreservedCellElements);
                WritePreservedElements(writer, shape.PreservedNonGeometrySections);
                WriteUserSection(writer, ns, shape.UserCells);
                WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
                WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
                WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
                WriteRawMasterInstanceChildShapes(writer, ns, shape, master, persistedIds);
                return;
            }

            if (WriteMasterDeltasOnly) {
                double masterWidth = master.Shape.Width > 0 ? master.Shape.Width : 1;
                double masterHeight = master.Shape.Height > 0 ? master.Shape.Height : 1;
                bool hasWidth = shape.Width > 0;
                bool hasHeight = shape.Height > 0;
                bool sizeDiffers = (hasWidth && Math.Abs(shape.Width - masterWidth) > 1e-12) ||
                                   (hasHeight && Math.Abs(shape.Height - masterHeight) > 1e-12);

                if (sizeDiffers) {
                    double width = hasWidth ? shape.Width : masterWidth;
                    double height = hasHeight ? shape.Height : masterHeight;
                    double locPinX = Math.Abs(shape.LocPinX) < double.Epsilon ? width / 2 : shape.LocPinX;
                    double locPinY = Math.Abs(shape.LocPinY) < double.Epsilon ? height / 2 : shape.LocPinY;
                    WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, locPinX, locPinY, shape.Angle);
                } else {
                    WriteCell(writer, ns, "PinX", shape.PinX);
                    WriteCell(writer, ns, "PinY", shape.PinY);

                    double masterLocPinX = Math.Abs(master.Shape.LocPinX) > double.Epsilon ? master.Shape.LocPinX : masterWidth / 2;
                    double masterLocPinY = Math.Abs(master.Shape.LocPinY) > double.Epsilon ? master.Shape.LocPinY : masterHeight / 2;
                    if (Math.Abs(shape.LocPinX) > double.Epsilon && Math.Abs(shape.LocPinX - masterLocPinX) > 1e-12) {
                        WriteCell(writer, ns, "LocPinX", shape.LocPinX);
                    }
                    if (Math.Abs(shape.LocPinY) > double.Epsilon && Math.Abs(shape.LocPinY - masterLocPinY) > 1e-12) {
                        WriteCell(writer, ns, "LocPinY", shape.LocPinY);
                    }
                    if (Math.Abs(shape.Angle - master.Shape.Angle) > 1e-12) {
                        WriteCell(writer, ns, "Angle", shape.Angle);
                    }
                }

            WriteCell(writer, ns, "ObjType", 1);
            WriteLayerMemberCell(writer, ns, shape.LayerNames, layerIndexes);
            WriteRelationshipCell(writer, ns, shape, persistedIds);
            WriteShapeLayoutCells(writer, ns, shape);
            WriteProtectionCells(writer, ns, shape.Protection);
            WriteTextBlockCells(writer, ns, shape.TextStyle);
            WriteCell(writer, ns, "LineWeight", shape.LineWeight);
                WriteCell(writer, ns, "LinePattern", shape.LinePattern);
                WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
                WriteCell(writer, ns, "FillPattern", shape.FillPattern);
                WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
                WritePreservedElements(writer, shape.PreservedCellElements);
                WritePreservedElements(writer, shape.PreservedNonGeometrySections);
                WriteUserSection(writer, ns, shape.UserCells);
                WriteTextStyleSections(writer, ns, shape.TextStyle);
                WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
                double geometryWidth = shape.Width > 0 ? shape.Width : masterWidth;
                double geometryHeight = shape.Height > 0 ? shape.Height : masterHeight;
                WriteShapeGeometry(writer, ns, shape.PreservedGeometrySections, master.NameU, geometryWidth, geometryHeight, writeGeneratedGeometryWhenEmpty: false);
                WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
                WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
                return;
            }

            double widthValue = shape.Width;
            if (widthValue <= 0 && master.Shape.Width > 0) {
                widthValue = master.Shape.Width;
            }
            if (widthValue <= 0) {
                widthValue = 1;
            }

            double heightValue = shape.Height;
            if (heightValue <= 0 && master.Shape.Height > 0) {
                heightValue = master.Shape.Height;
            }
            if (heightValue <= 0) {
                heightValue = 1;
            }

            double locPinXValue = Math.Abs(shape.LocPinX) < double.Epsilon ? widthValue / 2 : shape.LocPinX;
            double locPinYValue = Math.Abs(shape.LocPinY) < double.Epsilon ? heightValue / 2 : shape.LocPinY;

            WriteXForm(writer, ns, shape.PinX, shape.PinY, widthValue, heightValue, locPinXValue, locPinYValue, shape.Angle);
            WriteCell(writer, ns, "LineWeight", shape.LineWeight);
            WriteCell(writer, ns, "LinePattern", shape.LinePattern);
            WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
            WriteCell(writer, ns, "FillPattern", shape.FillPattern);
            WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
            WriteLayerMemberCell(writer, ns, shape.LayerNames, layerIndexes);
            WriteRelationshipCell(writer, ns, shape, persistedIds);
            WriteShapeLayoutCells(writer, ns, shape);
            WriteProtectionCells(writer, ns, shape.Protection);
            WriteTextBlockCells(writer, ns, shape.TextStyle);
            WritePreservedElements(writer, shape.PreservedCellElements);
            WritePreservedElements(writer, shape.PreservedNonGeometrySections);
            WriteUserSection(writer, ns, shape.UserCells);
            WriteTextStyleSections(writer, ns, shape.TextStyle);
            WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
            WriteShapeGeometry(writer, ns, shape.PreservedGeometrySections, master.NameU, widthValue, heightValue);
            WriteConnectionSection(writer, ns, shape.ConnectionPoints);
            WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
            WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
        }

        private static void WriteStandaloneShapeBody(XmlWriter writer, string ns, VisioShape shape, bool isGroup, KeyValuePair<string, string>? originalIdEntry, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyDictionary<string, int> layerIndexes) {
            double width = shape.Width > 0 ? shape.Width : 1;
            double height = shape.Height > 0 ? shape.Height : 1;
            double locPinX = Math.Abs(shape.LocPinX) < double.Epsilon ? width / 2 : shape.LocPinX;
            double locPinY = Math.Abs(shape.LocPinY) < double.Epsilon ? height / 2 : shape.LocPinY;

            WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, locPinX, locPinY, shape.Angle);
            WriteCell(writer, ns, "LineWeight", shape.LineWeight);
            WriteCell(writer, ns, "LinePattern", shape.LinePattern);
            WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
            WriteCell(writer, ns, "FillPattern", shape.FillPattern);
            WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
            if (!isGroup) {
                WriteCell(writer, ns, "ObjType", 1);
            }
            WriteLayerMemberCell(writer, ns, shape.LayerNames, layerIndexes);
            WriteRelationshipCell(writer, ns, shape, persistedIds);
            WriteShapeLayoutCells(writer, ns, shape);
            WriteProtectionCells(writer, ns, shape.Protection);
            WriteTextBlockCells(writer, ns, shape.TextStyle);
            WritePreservedElements(writer, shape.PreservedCellElements);
            WritePreservedElements(writer, shape.PreservedNonGeometrySections);
            WriteUserSection(writer, ns, shape.UserCells);
            WriteTextStyleSections(writer, ns, shape.TextStyle);
            WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
            WriteShapeGeometry(writer, ns, shape.PreservedGeometrySections, shape.NameU, width, height, writeGeneratedGeometryWhenEmpty: !isGroup);
            WriteConnectionSection(writer, ns, shape.ConnectionPoints);
            WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
            WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
        }

        private bool WriteStandaloneShapeBodyWithPreservedChildOrder(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            bool isGroup,
            KeyValuePair<string, string>? originalIdEntry,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, VisioMaster> effectiveMasters,
            IReadOnlyList<PackageMasterEntry> packageMasters,
            IReadOnlyDictionary<string, int> layerIndexes) {
            double width = shape.Width > 0 ? shape.Width : 1;
            double height = shape.Height > 0 ? shape.Height : 1;
            double locPinX = Math.Abs(shape.LocPinX) < double.Epsilon ? width / 2 : shape.LocPinX;
            double locPinY = Math.Abs(shape.LocPinY) < double.Epsilon ? height / 2 : shape.LocPinY;

            HashSet<string> emittedTokens = new(StringComparer.OrdinalIgnoreCase);
            bool wroteChildShapes = false;
            foreach (VisioShape.PreservedShapeChildEntry entry in shape.PreservedShapeChildren) {
                if (entry.RawElement != null) {
                    entry.RawElement.WriteTo(writer);
                    continue;
                }

                if (entry.Token is string token &&
                    !string.IsNullOrWhiteSpace(token) &&
                    emittedTokens.Add(token) &&
                    TryWriteStandaloneShapeChildToken(writer, ns, shape, isGroup, originalIdEntry, token, width, height, locPinX, locPinY, persistedIds, effectiveMasters, packageMasters, layerIndexes, ref wroteChildShapes)) {
                    if (string.Equals(token, "XForm", StringComparison.OrdinalIgnoreCase)) {
                        MarkShapeTransformCellTokens(emittedTokens);
                    }
                    continue;
                }
            }

            WriteRemainingStandaloneShapeChildren(writer, ns, shape, isGroup, originalIdEntry, emittedTokens, width, height, locPinX, locPinY, persistedIds, effectiveMasters, packageMasters, layerIndexes, ref wroteChildShapes);
            return wroteChildShapes;
        }

        private bool TryWriteStandaloneShapeChildToken(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            bool isGroup,
            KeyValuePair<string, string>? originalIdEntry,
            string token,
            double width,
            double height,
            double locPinX,
            double locPinY,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, VisioMaster> effectiveMasters,
            IReadOnlyList<PackageMasterEntry> packageMasters,
            IReadOnlyDictionary<string, int> layerIndexes,
            ref bool wroteChildShapes) {
            if (string.Equals(token, "XForm", StringComparison.OrdinalIgnoreCase)) {
                WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, locPinX, locPinY, shape.Angle);
                return true;
            }

            if (token.StartsWith("Cell:", StringComparison.OrdinalIgnoreCase)) {
                return TryWriteModeledShapeCell(writer, ns, shape, token.Substring("Cell:".Length), isGroup, width, height, locPinX, locPinY, persistedIds, layerIndexes);
            }

            if (string.Equals(token, "Section:User", StringComparison.OrdinalIgnoreCase)) {
                WriteUserSection(writer, ns, shape.UserCells);
                return true;
            }

            if (string.Equals(token, "Section:Char", StringComparison.OrdinalIgnoreCase)) {
                WriteCharSection(writer, ns, shape.TextStyle);
                return true;
            }

            if (string.Equals(token, "Section:Para", StringComparison.OrdinalIgnoreCase)) {
                WriteParaSection(writer, ns, shape.TextStyle);
                return true;
            }

            if (string.Equals(token, "Section:Geometry", StringComparison.OrdinalIgnoreCase)) {
                WriteShapeGeometry(writer, ns, shape.PreservedGeometrySections, shape.NameU, width, height, writeGeneratedGeometryWhenEmpty: !isGroup);
                return true;
            }

            if (string.Equals(token, "Section:Hyperlink", StringComparison.OrdinalIgnoreCase)) {
                WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
                return true;
            }

            if (string.Equals(token, "Section:Connection", StringComparison.OrdinalIgnoreCase)) {
                WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                return true;
            }

            if (string.Equals(token, "Section:Prop", StringComparison.OrdinalIgnoreCase)) {
                WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
                return true;
            }

            if (string.Equals(token, "Text", StringComparison.OrdinalIgnoreCase)) {
                WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
                return true;
            }

            if (string.Equals(token, "Shapes", StringComparison.OrdinalIgnoreCase)) {
                if (isGroup && shape.Children.Count > 0) {
                    WriteChildShapesContainer(writer, ns, shape, persistedIds, effectiveMasters, packageMasters, layerIndexes);
                }
                wroteChildShapes = true;
                return true;
            }

            return false;
        }

        private void WriteRemainingStandaloneShapeChildren(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            bool isGroup,
            KeyValuePair<string, string>? originalIdEntry,
            ISet<string> emittedTokens,
            double width,
            double height,
            double locPinX,
            double locPinY,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, VisioMaster> effectiveMasters,
            IReadOnlyList<PackageMasterEntry> packageMasters,
            IReadOnlyDictionary<string, int> layerIndexes,
            ref bool wroteChildShapes) {
            if (!HasAnyShapeTransformCellTokens(emittedTokens) && emittedTokens.Add("XForm")) {
                WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, locPinX, locPinY, shape.Angle);
                MarkShapeTransformCellTokens(emittedTokens);
            }

            WriteRemainingModeledShapeCells(writer, ns, shape, isGroup, width, height, locPinX, locPinY, emittedTokens, persistedIds, layerIndexes);

            if (emittedTokens.Add("Section:User")) {
                WriteUserSection(writer, ns, shape.UserCells);
            }

            if (emittedTokens.Add("Section:Char")) {
                WriteCharSection(writer, ns, shape.TextStyle);
            }

            if (emittedTokens.Add("Section:Para")) {
                WriteParaSection(writer, ns, shape.TextStyle);
            }

            if (emittedTokens.Add("Section:Hyperlink")) {
                WriteHyperlinkSection(writer, ns, shape.Hyperlinks);
            }

            if (emittedTokens.Add("Section:Geometry")) {
                WriteShapeGeometry(writer, ns, shape.PreservedGeometrySections, shape.NameU, width, height, writeGeneratedGeometryWhenEmpty: !isGroup);
            }

            if (emittedTokens.Add("Section:Connection")) {
                WriteConnectionSection(writer, ns, shape.ConnectionPoints);
            }

            if (emittedTokens.Add("Section:Prop")) {
                WriteDataSection(writer, ns, shape.Data, shape.PreservedDataRows, originalIdEntry, shape.ShapeData);
            }

            if (emittedTokens.Add("Text")) {
                WriteTextElement(writer, ns, shape.Text, shape.PreservedTextElement, shape.PreservedTextValue);
            }

            if (!wroteChildShapes && emittedTokens.Add("Shapes") && isGroup && shape.Children.Count > 0) {
                WriteChildShapesContainer(writer, ns, shape, persistedIds, effectiveMasters, packageMasters, layerIndexes);
                wroteChildShapes = true;
            }
        }

        private void WriteChildShapesContainer(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, VisioMaster> effectiveMasters,
            IReadOnlyList<PackageMasterEntry> packageMasters,
            IReadOnlyDictionary<string, int> layerIndexes) {
            writer.WriteStartElement("Shapes", ns);
            foreach (VisioShape child in shape.Children) {
                WriteShapeElement(writer, ns, child, persistedIds, effectiveMasters, packageMasters, layerIndexes);
            }
            writer.WriteEndElement();
        }
    }
}
