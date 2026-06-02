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
    /// <summary>
    /// Save core implementation for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>
        /// Core save routine that writes the VSdx structure.
        /// </summary>
        /// <param name="filePath">Target path.</param>
        private void SaveInternalCore(string filePath) {
            bool includeTheme = Theme != null;
            List<VisioPage> pagesToSave = _pages.Count > 0 ? _pages : new List<VisioPage> { new VisioPage("Page-1") { Id = 0 } };
            PrepareTextFontFaceNames(pagesToSave);
            int pageCount = pagesToSave.Count;
            List<string> pagePartNames = new();
            int masterCount;

            using (Package package = Package.Open(filePath, FileMode.Create)) {
                masterCount = WritePackage(package, includeTheme, pagesToSave, pageCount, pagePartNames);
            }

            FixContentTypes(filePath, masterCount, includeTheme, pagePartNames);
        }

        /// <summary>
        /// Core save routine that writes the VSDX structure to a stream.
        /// </summary>
        /// <param name="destination">Target stream.</param>
        private void SaveInternalCore(Stream destination) {
            bool includeTheme = Theme != null;
            List<VisioPage> pagesToSave = _pages.Count > 0 ? _pages : new List<VisioPage> { new VisioPage("Page-1") { Id = 0 } };
            PrepareTextFontFaceNames(pagesToSave);
            int pageCount = pagesToSave.Count;
            List<string> pagePartNames = new();
            int masterCount;

            using var packageStream = new MemoryStream();
            using (Package package = Package.Open(packageStream, FileMode.Create, FileAccess.ReadWrite)) {
                masterCount = WritePackage(package, includeTheme, pagesToSave, pageCount, pagePartNames);
            }

            FixContentTypes(packageStream, masterCount, includeTheme, pagePartNames);

            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
                destination.SetLength(0);
            }
            packageStream.Seek(0, SeekOrigin.Begin);
            packageStream.CopyTo(destination);
            destination.Flush();
            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
            }
        }

        private int WritePackage(
            Package package,
            bool includeTheme,
            List<VisioPage> pagesToSave,
            int pageCount,
            List<string> pagePartNames) {
            int masterCount;
            Uri documentUri = new("/visio/document.xml", UriKind.Relative);
            PackagePart documentPart = package.CreatePart(documentUri, DocumentContentType);
            package.CreateRelationship(documentUri, TargetMode.Internal, DocumentRelationshipType, "rId1");

                Uri coreUri = new("/docProps/core.xml", UriKind.Relative);
                PackagePart corePart = package.CreatePart(coreUri, "application/vnd.openxmlformats-package.core-properties+xml");
                package.CreateRelationship(coreUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId2");

                Uri appUri = new("/docProps/app.xml", UriKind.Relative);
                PackagePart appPart = package.CreatePart(appUri, "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                package.CreateRelationship(appUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId3");

                Uri customUri = new("/docProps/custom.xml", UriKind.Relative);
                PackagePart customPart = package.CreatePart(customUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml");
                package.CreateRelationship(customUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "rId4");

                Uri thumbUri = new("/docProps/thumbnail.emf", UriKind.Relative);
                PackagePart thumbPart = package.CreatePart(thumbUri, "image/x-emf");
                package.CreateRelationship(thumbUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail", "rId5");

                Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
                PackagePart pagesPart = package.CreatePart(pagesUri, PagesContentType);
                documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, PagesRelationshipType, "rId1");

                Uri windowsUri = new("/visio/windows.xml", UriKind.Relative);
                PackagePart windowsPart = package.CreatePart(windowsUri, WindowsContentType);
                documentPart.CreateRelationship(new Uri("windows.xml", UriKind.Relative), TargetMode.Internal, WindowsRelationshipType, "rId2");

                PackagePart? themePart = null;
                if (includeTheme) {
                    Uri themeUri = new("/visio/theme/theme1.xml", UriKind.Relative);
                    themePart = package.CreatePart(themeUri, ThemeContentType);
                    documentPart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal, ThemeRelationshipType, "rId3");
                }

                List<(VisioPage Page, PackagePart Part, PackageRelationship Relationship)> pageParts = new();
                for (int i = 0; i < pagesToSave.Count; i++) {
                    VisioPage currentPage = pagesToSave[i];
                    Uri pageUri = new($"/visio/pages/page{i + 1}.xml", UriKind.Relative);
                    PackagePart pagePart = package.CreatePart(pageUri, PageContentType);
                    PackageRelationship pageRelationship = pagesPart.CreateRelationship(new Uri($"page{i + 1}.xml", UriKind.Relative), TargetMode.Internal, PageRelationshipType, $"rId{i + 1}");
                    pageParts.Add((currentPage, pagePart, pageRelationship));
                }

                XmlWriterSettings settings = new() {
                    Encoding = new UTF8Encoding(false),
                    CloseOutput = true,
                    Indent = false,
                };
                using (XmlWriter writer = XmlWriter.Create(corePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                    writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
                    writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
                    writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
                    writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                if (!string.IsNullOrEmpty(Title)) {
                    package.PackageProperties.Title = Title;
                }
                if (!string.IsNullOrEmpty(Author)) {
                    package.PackageProperties.Creator = Author;
                }

                const string ns = VisioNamespace;

                if (themePart != null && Theme != null) {
                    if (Theme.TemplateXml != null) {
                        XDocument themeXml = new(Theme.TemplateXml);
                        if (themeXml.Root != null) {
                            themeXml.Root.SetAttributeValue("name", Theme.Name);
                        }

                        using Stream s = themePart.GetStream(FileMode.Create, FileAccess.Write);
                        using StreamWriter sw = new(s, new UTF8Encoding(false));
                        sw.Write(themeXml.Declaration + Environment.NewLine + themeXml.ToString(SaveOptions.DisableFormatting));
                    } else {
                        using (XmlWriter writer = XmlWriter.Create(themePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("a", "theme", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            if (!string.IsNullOrEmpty(Theme.Name)) {
                                writer.WriteAttributeString("name", Theme.Name);
                            }
                            writer.WriteEndElement();
                            writer.WriteEndDocument();
                        }
                    }
                }

                // Write visio/document.xml
                {
                    XDocument docXml = CreateVisioDocumentXml(
                        _requestRecalcOnOpen,
                        PreservedDocumentAttributes,
                        PreservedDocumentElements,
                        PreservedDocumentSettingsAttributes,
                        PreservedDocumentSettingsElements,
                        PreservedColorsAttributes,
                        PreservedColorsElements,
                        PreservedFaceNamesAttributes,
                        PreservedFaceNamesElements,
                        PreservedStyleSheetsAttributes,
                        PreservedStyleSheetsElements,
                        PreservedGeneratedStyleSheets,
                        PreservedAdditionalStyleSheets);
                    using Stream s = documentPart.GetStream(FileMode.Create, FileAccess.Write);
                    using StreamWriter sw = new(s, new UTF8Encoding(false));
                    sw.Write(docXml.Declaration + Environment.NewLine + docXml.ToString(SaveOptions.DisableFormatting));
                }

                // Write visio/windows.xml with minimal expected structure
                {
                    XNamespace vNs = VisioNamespace;
                    XElement root = new(vNs + "Windows",
                        new XAttribute("ClientWidth", XmlConvert.ToString(8.5)),
                        new XAttribute("ClientHeight", XmlConvert.ToString(11.0)));
                    root.Add(new XElement(vNs + "Window",
                        new XAttribute("WindowType", 1),
                        new XAttribute("WindowState", 0),
                        new XAttribute("ClientWidth", XmlConvert.ToString(8.5)),
                        new XAttribute("ClientHeight", XmlConvert.ToString(11.0))));
                    XDocument winXml = new(root);
                    using Stream s = windowsPart.GetStream(FileMode.Create, FileAccess.Write);
                    using StreamWriter sw = new(s, new UTF8Encoding(false));
                    sw.Write(winXml.Declaration + Environment.NewLine + winXml.ToString(SaveOptions.DisableFormatting));
                }

                ValidatePagesForSave(pagesToSave);

                Dictionary<VisioPage, Dictionary<string, VisioMaster>> effectivePageMasters = new();
                List<VisioMaster> masterCandidates = new();
                foreach (VisioPage page in pagesToSave) {
                    Dictionary<string, VisioMaster> pageMasters = BuildEffectiveShapeMasterMap(page);
                    effectivePageMasters[page] = pageMasters;
                    AddMastersInShapeOrder(page.Shapes, pageMasters, masterCandidates);

                    if (UseMastersByDefault) {
                        foreach (VisioConnector connector in page.Connectors) {
                            if (connector.Kind == ConnectorKind.Dynamic) {
                                masterCandidates.Add(EnsureBuiltinMaster("Dynamic connector"));
                            }
                        }
                    }
                }

                List<PackageMasterEntry> masters = CreatePackageMasterEntries(masterCandidates);

                PackagePart? mastersPart = null;
                if (masters.Count > 0) {
                    Uri mastersUri = new("/visio/masters/masters.xml", UriKind.Relative);
                    mastersPart = package.CreatePart(mastersUri, "application/vnd.ms-visio.masters+xml");
                    documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative), TargetMode.Internal, MastersRelationshipType, "rId4");

                    for (int i = 0; i < masters.Count; i++) {
                        PackageMasterEntry entry = masters[i];
                        VisioMaster master = entry.Master;
                        Uri masterUri = new($"/visio/masters/master{entry.PartNumber}.xml", UriKind.Relative);
                        PackagePart masterPart = package.CreatePart(masterUri, "application/vnd.ms-visio.master+xml");
                        mastersPart.CreateRelationship(new Uri($"master{entry.PartNumber}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{entry.PartNumber}");
                        foreach ((_, PackagePart part, _) in pageParts) {
                            part.CreateRelationship(new Uri($"../masters/master{entry.PartNumber}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{entry.PartNumber}");
                        }

                        if (master.RawMasterContentXml != null) {
                            using Stream stream = masterPart.GetStream(FileMode.Create, FileAccess.Write);
                            using StreamWriter streamWriter = new(stream, new UTF8Encoding(false));
                            XDocument rawMasterContent = MergeRawMasterMetadata(master.RawMasterContentXml, master);
                            streamWriter.Write(rawMasterContent.Declaration + Environment.NewLine + rawMasterContent.ToString(SaveOptions.DisableFormatting));
                            WriteRawMasterRelationships(package, masterPart, master, entry.PartNumber);
                            continue;
                        }

                        using (XmlWriter writer = XmlWriter.Create(masterPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("MasterContents", ns);
                            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                            WritePreservedAttributes(writer, master.PreservedMasterContentAttributes);
                            writer.WriteStartElement("Shapes", ns);
                            WritePreservedAttributes(writer, master.PreservedShapesAttributes);
                            VisioShape s = master.Shape;
                            double masterWidth = s.Width > 0 ? s.Width : 1;
                            double masterHeight = s.Height > 0 ? s.Height : 1;
                            double masterLocPinX = Math.Abs(s.LocPinX) < double.Epsilon ? masterWidth / 2 : s.LocPinX;
                            double masterLocPinY = Math.Abs(s.LocPinY) < double.Epsilon ? masterHeight / 2 : s.LocPinY;
                            TryGetBuiltinMasterDefinition(master.NameU, out var masterDefinition);
                            writer.WriteStartElement("Shape", ns);
                            writer.WriteAttributeString("ID", "1");
                            string masterShapeName = s.Name ?? s.NameU ?? "MasterShape";
                            writer.WriteAttributeString("Name", masterShapeName);
                            writer.WriteAttributeString("NameU", master.NameU);
                            writer.WriteAttributeString("Type", "Shape");
                            if (masterDefinition?.GeometryKind == BuiltinGeometryKind.DynamicConnector) {
                                writer.WriteAttributeString("LineStyle", "0");
                                writer.WriteAttributeString("FillStyle", "0");
                                writer.WriteAttributeString("TextStyle", "0");
                                WriteXForm1D(writer, ns, 0, 0, 1, 0);
                                WriteCell(writer, ns, "OneD", 1);
                                WriteCell(writer, ns, "ObjType", 2);
                                WriteCell(writer, ns, "LineWeight", s.LineWeight);
                                WriteCell(writer, ns, "LinePattern", s.LinePattern);
                                WriteCellValue(writer, ns, "LineColor", s.LineColor.ToVisioHex());
                                WriteCell(writer, ns, "FillPattern", 0);
                                WriteCellValue(writer, ns, "FillForegnd", Color.Transparent.ToVisioHex());
                                WriteCell(writer, ns, "LockHeight", 1);
                                WriteCell(writer, ns, "LockCalcWH", 1);
                                WriteCell(writer, ns, "GlueType", 2);
                                WriteCell(writer, ns, "NoAlignBox", 1);
                                WriteCell(writer, ns, "DynFeedback", 2);
                                WriteCell(writer, ns, "ShapeSplittable", 1);
                                WriteCell(writer, ns, "LayerMember", 0);
                                WriteConnectorControlSection(writer, ns, masterHeight);
                            } else {
                                writer.WriteAttributeString("LineStyle", "1");
                                writer.WriteAttributeString("FillStyle", "1");
                                writer.WriteAttributeString("TextStyle", "1");
                                WriteXForm(writer, ns, s.PinX, s.PinY, masterWidth, masterHeight, masterLocPinX, masterLocPinY, s.Angle);
                                WriteCell(writer, ns, "ObjType", 1);
                                if (masterDefinition?.LockAspect == true) {
                                    WriteCell(writer, ns, "LockAspect", 1);
                                }
                                WriteCell(writer, ns, "LineWeight", s.LineWeight);
                                WriteCell(writer, ns, "LinePattern", s.LinePattern);
                                WriteCellValue(writer, ns, "LineColor", s.LineColor.ToVisioHex());
                                WriteCell(writer, ns, "FillPattern", s.FillPattern);
                                WriteCellValue(writer, ns, "FillForegnd", s.FillColor.ToVisioHex());
                                WriteShapeGeometry(writer, ns, s.PreservedGeometrySections, master.NameU, masterWidth, masterHeight);
                                WriteDefaultTextBlock(writer, ns, masterWidth, masterHeight);
                            }
                            WriteCell(writer, ns, "ShapeSplit", 1);
                            WriteCell(writer, ns, "QuickStyleType", 2);
                            WriteConnectionSection(writer, ns, s.ConnectionPoints);
                            WriteMasterUserSection(writer, ns);
                            WriteMasterCharacterSection(writer, ns);
                            WriteDataSection(writer, ns, s.Data, s.PreservedDataRows, shapeDataRows: s.ShapeData);
                            WriteTextElement(writer, ns, s.Text, s.PreservedTextElement, s.PreservedTextValue);
                            writer.WriteEndElement();
                            WritePreservedElements(writer, master.PreservedAdditionalShapeElements);
                            writer.WriteEndElement();
                            WritePreservedElements(writer, master.PreservedMasterContentElements);
                            writer.WriteEndElement();
                            writer.WriteEndDocument();
                        }
                    }

                    // Write masters list (masters.xml)
                    using (XmlWriter writer = XmlWriter.Create(mastersPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Masters", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        VisioMaster mastersRootMetadataSource = GetMastersRootMetadataSource(masters);
                        WritePreservedAttributes(writer, mastersRootMetadataSource.PreservedMastersRootAttributes);
                        WritePreservedElements(writer, mastersRootMetadataSource.PreservedMastersRootElements);
                        for (int i = 0; i < masters.Count; i++) {
                            PackageMasterEntry entry = masters[i];
                            VisioMaster m = entry.Master;
                            TryGetBuiltinMasterDefinition(m.NameU, out var masterDefinition);
                            writer.WriteStartElement("Master", ns);
                            writer.WriteAttributeString("ID", entry.PackageId);
                            writer.WriteAttributeString("Name", m.NameU);
                            writer.WriteAttributeString("NameU", m.NameU);
                            writer.WriteAttributeString("IsCustomNameU", "1");
                            writer.WriteAttributeString("IsCustomName", "1");
                            writer.WriteAttributeString("Prompt", masterDefinition?.Prompt ?? "Drag onto the page.");
                            writer.WriteAttributeString("IconSize", "1");
                            writer.WriteAttributeString("AlignName", "2");
                            writer.WriteAttributeString("MatchByName", masterDefinition?.MatchByName == true ? "1" : "0");
                            writer.WriteAttributeString("IconUpdate", masterDefinition?.IconUpdate == false ? "0" : "1");
                            writer.WriteAttributeString("UniqueID", masterDefinition?.UniqueId ?? Guid.Empty.ToString("B").ToUpperInvariant());
                            writer.WriteAttributeString("BaseID", masterDefinition?.BaseId ?? Guid.Empty.ToString("B").ToUpperInvariant());
                            writer.WriteAttributeString("PatternFlags", "0");
                            writer.WriteAttributeString("Hidden", "0");
                            writer.WriteAttributeString("MasterType", XmlConvert.ToString(masterDefinition?.MasterType ?? 2));
                            foreach (XAttribute preservedAttribute in m.PreservedMasterAttributes) {
                                writer.WriteAttributeString(
                                    preservedAttribute.Name.LocalName,
                                    preservedAttribute.Name.NamespaceName.Length == 0 ? null : preservedAttribute.Name.NamespaceName,
                                    preservedAttribute.Value);
                            }
                            WriteMasterPageSheet(writer, ns, m, masterDefinition);
                            WritePreservedElements(writer, m.PreservedMasterElements);
                            writer.WriteStartElement("Rel", ns);
                            writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{entry.PartNumber}");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }

                using (XmlWriter writer = XmlWriter.Create(pagesPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Pages", ns);
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                    for (int i = 0; i < pageParts.Count; i++) {
                        (VisioPage page, _, PackageRelationship pageRelationship) = pageParts[i];
                        writer.WriteStartElement("Page", ns);
                        writer.WriteAttributeString("ID", XmlConvert.ToString(page.Id));
                        writer.WriteAttributeString("Name", page.Name);
                        writer.WriteAttributeString("NameU", page.NameU ?? page.Name);
                        if (page.IsBackground) {
                            writer.WriteAttributeString("Background", "1");
                        }

                        int? backgroundPageId = page.BackgroundPage?.Id ?? page.BackgroundPageId;
                        if (backgroundPageId.HasValue) {
                            writer.WriteAttributeString("BackPage", XmlConvert.ToString(backgroundPageId.Value));
                        }

                        double viewScale = page.ViewScale;
                        if (double.IsNaN(viewScale) || double.IsInfinity(viewScale) || viewScale <= 0) {
                            viewScale = 1;
                        }
                        writer.WriteAttributeString("ViewScale", XmlConvert.ToString(viewScale));
                        writer.WriteAttributeString("ViewCenterX", XmlConvert.ToString(page.ViewCenterX));
                        writer.WriteAttributeString("ViewCenterY", XmlConvert.ToString(page.ViewCenterY));
                        foreach (XAttribute preservedAttribute in page.PreservedPageAttributes) {
                            writer.WriteAttributeString(
                                preservedAttribute.Name.LocalName,
                                preservedAttribute.Name.NamespaceName.Length == 0 ? null : preservedAttribute.Name.NamespaceName,
                                preservedAttribute.Value);
                        }

                        writer.WriteStartElement("PageSheet", ns);
                        writer.WriteAttributeString("LineStyle", "0");
                        writer.WriteAttributeString("FillStyle", "0");
                        writer.WriteAttributeString("TextStyle", "0");

                        bool useUnits = page.DefaultUnit != VisioMeasurementUnit.Inches ||
                                        page.Width != 8.26771653543307 ||
                                        page.Height != 11.69291338582677;
                        if (useUnits) {
                            string pageUnitCode = page.DefaultUnit.ToVisioUnitCode();
                            WritePageCell(writer, ns, "PageWidth", page.Width.FromInches(page.DefaultUnit), pageUnitCode);
                            WritePageCell(writer, ns, "PageHeight", page.Height.FromInches(page.DefaultUnit), pageUnitCode);
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
                        writer.WriteStartElement("Rel", ns);
                        writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", pageRelationship.Id);
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }

                foreach ((VisioPage page, PackagePart pagePart, _) in pageParts) {
                    using (XmlWriter writer = XmlWriter.Create(pagePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        Dictionary<string, VisioMaster> pageMasters = effectivePageMasters[page];
                        Dictionary<string, string> persistedIds = BuildPersistedIdMap(page, pageMasters);
                        writer.WriteStartDocument();
                        writer.WriteStartElement("PageContents", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        writer.WriteAttributeString("xml", "space", null, "preserve");
                        WritePreservedAttributes(writer, page.PreservedPageContentAttributes);
                        WritePreservedElements(writer, page.PreservedPageContentElements);
                        Dictionary<string, int> layerIndexes = BuildLayerIndexMap(page, out _);
                        bool writeShapesContainer = page.Shapes.Count > 0 ||
                                                    page.Connectors.Count > 0 ||
                                                    page.PreservedShapesContainerAttributes.Count > 0 ||
                                                    page.PreservedShapesContainerElements.Count > 0 ||
                                                    page.PreservedShapesChildren.Count > 0;
                        if (writeShapesContainer) {
                            writer.WriteStartElement("Shapes", ns);
                            WritePreservedAttributes(writer, page.PreservedShapesContainerAttributes);
                            HashSet<VisioShape> emittedShapes = new();
                            HashSet<VisioConnector> emittedConnectors = new();
                            List<VisioShape> currentShapes = page.Shapes.ToList();
                            List<VisioConnector> currentConnectors = page.Connectors.ToList();
                            int nextShapeIndex = 0;
                            int nextConnectorIndex = 0;
                            if (page.PreservedShapesChildren.Count > 0) {
                                foreach (VisioPage.PreservedShapeChildEntry entry in page.PreservedShapesChildren) {
                                    if (entry.RawElement != null) {
                                        entry.RawElement.WriteTo(writer);
                                        continue;
                                    }

                                    if (entry.Shape != null &&
                                        TryGetNextUnemittedShape(currentShapes, emittedShapes, ref nextShapeIndex, out VisioShape? shapeToEmit) &&
                                        shapeToEmit != null) {
                                        WriteShapeElement(writer, ns, shapeToEmit, persistedIds, pageMasters, masters, layerIndexes);
                                        emittedShapes.Add(shapeToEmit);
                                        continue;
                                    }

                                    if (entry.Connector != null &&
                                        TryGetNextUnemittedConnector(currentConnectors, emittedConnectors, ref nextConnectorIndex, out VisioConnector? connectorToEmit) &&
                                        connectorToEmit != null) {
                                        WriteConnectorShapeElement(writer, ns, connectorToEmit, persistedIds, masters, layerIndexes);
                                        emittedConnectors.Add(connectorToEmit);
                                    }
                                }
                            } else {
                                WritePreservedElements(writer, page.PreservedShapesContainerElements);
                            }

                            foreach (VisioShape shape in page.Shapes) {
                                if (emittedShapes.Add(shape)) {
                                    WriteShapeElement(writer, ns, shape, persistedIds, pageMasters, masters, layerIndexes);
                                }
                            }

                            foreach (VisioConnector connector in page.Connectors) {
                                if (emittedConnectors.Add(connector)) {
                                    WriteConnectorShapeElement(writer, ns, connector, persistedIds, masters, layerIndexes);
                                }
                            }

                            writer.WriteEndElement(); // Shapes

                        }

                        bool writeConnectsContainer = page.Connectors.Count > 0 ||
                                                      page.PreservedConnectsAttributes.Count > 0 ||
                                                      page.PreservedConnectsElements.Count > 0 ||
                                                      page.PreservedConnectRows.Count > 0;
                        if (writeConnectsContainer) {
                            writer.WriteStartElement("Connects", ns);
                            WritePreservedAttributes(writer, page.PreservedConnectsAttributes);
                            HashSet<(VisioConnector Connector, VisioConnectorEndpointScope Endpoint)> emittedConnectRows = new();
                            if (page.PreservedConnectChildren.Count > 0) {
                                foreach (VisioPage.PreservedConnectChildEntry entry in page.PreservedConnectChildren) {
                                    if (entry.RawElement != null) {
                                        entry.RawElement.WriteTo(writer);
                                        continue;
                                    }

                                    if (entry.Connector == null ||
                                        !page.Connectors.Contains(entry.Connector) ||
                                        entry.EndpointScope is not VisioConnectorEndpointScope.Start and not VisioConnectorEndpointScope.End) {
                                        continue;
                                    }

                                    WriteConnectElement(writer, ns, persistedIds, entry.Connector, entry.EndpointScope.Value);
                                    emittedConnectRows.Add((entry.Connector, entry.EndpointScope.Value));
                                }
                            } else if (page.PreservedConnectRows.Count > 0) {
                                WritePreservedElements(writer, page.PreservedConnectsElements);
                                foreach (VisioPage.PreservedConnectRowEntry entry in page.PreservedConnectRows) {
                                    if (entry.RawElement != null) {
                                        entry.RawElement.WriteTo(writer);
                                        continue;
                                    }

                                    if (entry.Connector == null ||
                                        !page.Connectors.Contains(entry.Connector) ||
                                        entry.EndpointScope is not VisioConnectorEndpointScope.Start and not VisioConnectorEndpointScope.End) {
                                        continue;
                                    }

                                    WriteConnectElement(writer, ns, persistedIds, entry.Connector, entry.EndpointScope.Value);
                                    emittedConnectRows.Add((entry.Connector, entry.EndpointScope.Value));
                                }
                            } else {
                                WritePreservedElements(writer, page.PreservedConnectsElements);
                            }

                            foreach (VisioConnector connector in page.Connectors) {
                                if (!emittedConnectRows.Contains((connector, VisioConnectorEndpointScope.Start))) {
                                    WriteConnectElement(writer, ns, persistedIds, connector, VisioConnectorEndpointScope.Start);
                                }

                                if (!emittedConnectRows.Contains((connector, VisioConnectorEndpointScope.End))) {
                                    WriteConnectElement(writer, ns, persistedIds, connector, VisioConnectorEndpointScope.End);
                                }
                            }
                            writer.WriteEndElement(); // Connects
                        }

                        writer.WriteEndElement(); // PageContents
                        writer.WriteEndDocument();
                    }
                }
                masterCount = masters.Count;
                pagePartNames.Clear();
                pagePartNames.AddRange(pageParts
                    .Select(part => part.Part.Uri.OriginalString)
                    .Distinct(StringComparer.OrdinalIgnoreCase));

            return masterCount;
        }

    }
}
