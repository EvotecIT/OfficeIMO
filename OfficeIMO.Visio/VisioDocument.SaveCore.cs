using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using SixLabors.ImageSharp;

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
                    XDocument docXml = CreateVisioDocumentXml(_requestRecalcOnOpen);
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

                        using (XmlWriter writer = XmlWriter.Create(masterPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("MasterContents", ns);
                            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
                            writer.WriteStartElement("Shapes", ns);
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
                            writer.WriteAttributeString("LineStyle", "0");
                            writer.WriteAttributeString("FillStyle", "0");
                            writer.WriteAttributeString("TextStyle", "0");
                            if (masterDefinition?.GeometryKind == BuiltinGeometryKind.DynamicConnector) {
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
                                WriteMasterGeometry(writer, ns, master.NameU, masterWidth, masterHeight);
                                WriteDefaultTextBlock(writer, ns, masterWidth, masterHeight);
                            }
                            WriteCell(writer, ns, "ShapeSplit", 1);
                            WriteCell(writer, ns, "QuickStyleType", 2);
                            WriteConnectionSection(writer, ns, s.ConnectionPoints);
                            WriteMasterUserSection(writer, ns);
                            WriteMasterCharacterSection(writer, ns);
                            WriteDataSection(writer, ns, s.Data);
                            WriteTextElement(writer, ns, s.Text);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndDocument();
                        }
                    }

                    // Write masters list (masters.xml)
                    using (XmlWriter writer = XmlWriter.Create(mastersPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Masters", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
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
                            WriteMasterPageSheet(writer, ns, masterDefinition);
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
                        writer.WriteAttributeString("NameU", page.Name);
                        double viewScale = page.ViewScale;
                        if (double.IsNaN(viewScale) || double.IsInfinity(viewScale) || viewScale <= 0) {
                            viewScale = 1;
                        }
                        writer.WriteAttributeString("ViewScale", XmlConvert.ToString(viewScale));
                        writer.WriteAttributeString("ViewCenterX", XmlConvert.ToString(page.ViewCenterX));
                        writer.WriteAttributeString("ViewCenterY", XmlConvert.ToString(page.ViewCenterY));

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
                        WritePageCell(writer, ns, "DrawingSizeType", 0);
                        WritePageCell(writer, ns, "DrawingScaleType", 0);
                        WritePageCell(writer, ns, "InhibitSnap", 0);
                        WritePageCell(writer, ns, "PageLockReplace", 0, "BOOL");
                        WritePageCell(writer, ns, "PageLockDuplicate", 0, "BOOL");
                        WritePageCell(writer, ns, "UIVisibility", 0);
                        WritePageCell(writer, ns, "ShdwType", 0);
                        WritePageCell(writer, ns, "ShdwObliqueAngle", 0);
                        WritePageCell(writer, ns, "ShdwScaleFactor", 1);
                        WritePageCell(writer, ns, "DrawingResizeType", 1);
                        WritePageCell(writer, ns, "PageShapeSplit", 1);
                        // For non-default page sizes, include theme/margin metadata like the asset samples
                        if (useUnits) {
                            WritePageCell(writer, ns, "ColorSchemeIndex", 60);
                            WritePageCell(writer, ns, "EffectSchemeIndex", 60);
                            WritePageCell(writer, ns, "ConnectorSchemeIndex", 60);
                            WritePageCell(writer, ns, "FontSchemeIndex", 60);
                            WritePageCell(writer, ns, "ThemeIndex", 60);
                            WritePageCell(writer, ns, "PageLeftMargin", 0.25, "MM");
                            WritePageCell(writer, ns, "PageRightMargin", 0.25, "MM");
                            WritePageCell(writer, ns, "PageTopMargin", 0.25, "MM");
                            WritePageCell(writer, ns, "PageBottomMargin", 0.25, "MM");
                            WritePageCell(writer, ns, "PrintPageOrientation", 2);
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
                        Dictionary<string, string> persistedIds = BuildPersistedIdMap(page);
                        Dictionary<string, VisioMaster> pageMasters = effectivePageMasters[page];
                        writer.WriteStartDocument();
                        writer.WriteStartElement("PageContents", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        writer.WriteAttributeString("xml", "space", null, "preserve");
                        if (page.Shapes.Count > 0 || page.Connectors.Count > 0) {
                            writer.WriteStartElement("Shapes", ns);

                            foreach (VisioShape shape in page.Shapes) {
                                WriteShapeElement(writer, ns, shape, persistedIds, pageMasters, masters);
                            }

                            foreach (VisioConnector connector in page.Connectors) {
                                writer.WriteStartElement("Shape", ns);
                                writer.WriteAttributeString("ID", GetPersistedId(persistedIds, connector.Id));
                                bool isDynamic = connector.Kind == ConnectorKind.Dynamic;
                                string connName = (isDynamic && UseMastersByDefault) ? "Dynamic connector" : "Connector";
                                writer.WriteAttributeString("Name", connName);
                                writer.WriteAttributeString("NameU", connName);
                                writer.WriteAttributeString("LineStyle", "0");
                                writer.WriteAttributeString("FillStyle", "0");
                                writer.WriteAttributeString("TextStyle", "0");
                                if (isDynamic && UseMastersByDefault) {
                                    var m = EnsureBuiltinMaster("Dynamic connector");
                                    writer.WriteAttributeString("Master", GetPackageMasterId(masters, m));
                                }
                                double startX, startY, endX, endY;
                                if (connector.FromConnectionPoint != null) {
                                    (startX, startY) = connector.From.GetAbsolutePoint(connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y);
                                } else {
                                    var (fL, fB, fR, fT) = connector.From.GetBounds();
                                    // Choose side based on relative centers to respect rotation/locpin.
                                    var (tL2, _, tR2, _) = connector.To.GetBounds();
                                    double fromCx = (fL + fR) / 2.0;
                                    double toCx = (tL2 + tR2) / 2.0;
                                    bool toIsRight = toCx >= fromCx;
                                    startX = toIsRight ? fR : fL;
                                    startY = (fB + fT) / 2.0;
                                }

                                if (connector.ToConnectionPoint != null) {
                                    (endX, endY) = connector.To.GetAbsolutePoint(connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y);
                                } else {
                                    var (tL, tB, tR, tT) = connector.To.GetBounds();
                                    // Choose side based on relative centers
                                    var (fL2, _, fR2, _) = connector.From.GetBounds();
                                    double toCx = (tL + tR) / 2.0;
                                    double fromCx = (fL2 + fR2) / 2.0;
                                    bool fromIsLeft = fromCx <= toCx;
                                    endX = fromIsLeft ? tL : tR;
                                    endY = (tB + tT) / 2.0;
                                }
                                WriteXForm1D(writer, ns, startX, startY, endX, endY);
                                WriteCell(writer, ns, "LineWeight", connector.LineWeight);
                                WriteCell(writer, ns, "LinePattern", connector.LinePattern);
                                WriteCellValue(writer, ns, "LineColor", connector.LineColor.ToVisioHex());
                                WriteCell(writer, ns, "FillPattern", 0);
                                WriteCellValue(writer, ns, "FillForegnd", Color.Transparent.ToVisioHex());
                                WriteCell(writer, ns, "OneD", 1);
                                if (connector.BeginArrow.HasValue) {
                                    WriteCell(writer, ns, "BeginArrow", (int)connector.BeginArrow.Value);
                                }
                                if (connector.EndArrow.HasValue) {
                                    WriteCell(writer, ns, "EndArrow", (int)connector.EndArrow.Value);
                                }

                                WriteConnectorGeometry(writer, ns, connector, startX, startY, endX, endY);

                                KeyValuePair<string, string>? connectorOriginalId = GetOriginalIdEntry(persistedIds, connector.Id);
                                WriteDataSection(writer, ns, new Dictionary<string, string>(), connectorOriginalId);
                                WriteTextElement(writer, ns, connector.Label);
                                writer.WriteEndElement();
                            }

                            writer.WriteEndElement(); // Shapes

                            if (page.Connectors.Count > 0) {
                                writer.WriteStartElement("Connects", ns);
                                foreach (VisioConnector connector in page.Connectors) {
                                    writer.WriteStartElement("Connect", ns);
                                    writer.WriteAttributeString("FromSheet", GetPersistedId(persistedIds, connector.Id));
                                    writer.WriteAttributeString("FromCell", "BeginX");
                                    writer.WriteAttributeString("ToSheet", GetPersistedId(persistedIds, connector.From.Id));
                                    writer.WriteAttributeString("ToCell", GetConnectionCell(connector.From, connector.FromConnectionPoint));
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("Connect", ns);
                                    writer.WriteAttributeString("FromSheet", GetPersistedId(persistedIds, connector.Id));
                                    writer.WriteAttributeString("FromCell", "EndX");
                                    writer.WriteAttributeString("ToSheet", GetPersistedId(persistedIds, connector.To.Id));
                                    writer.WriteAttributeString("ToCell", GetConnectionCell(connector.To, connector.ToConnectionPoint));
                                    writer.WriteEndElement();
                                }
                                writer.WriteEndElement(); // Connects
                            }
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

        private void WriteShapeElement(XmlWriter writer, string ns, VisioShape shape, IReadOnlyDictionary<string, string> persistedIds, IReadOnlyDictionary<string, VisioMaster> effectiveMasters, IReadOnlyList<PackageMasterEntry> packageMasters) {
            writer.WriteStartElement("Shape", ns);
            writer.WriteAttributeString("ID", GetPersistedId(persistedIds, shape.Id));
            string shapeName = shape.Name ?? shape.NameU ?? $"Shape{shape.Id}";
            writer.WriteAttributeString("Name", shapeName);
            VisioMaster? effectiveMaster = TryGetEffectiveMaster(effectiveMasters, shape);
            writer.WriteAttributeString("NameU", shape.NameU ?? effectiveMaster?.NameU ?? shapeName);

            bool isGroup = string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase) || shape.Children.Count > 0;
            writer.WriteAttributeString("Type", isGroup ? "Group" : "Shape");
            writer.WriteAttributeString("LineStyle", "0");
            writer.WriteAttributeString("FillStyle", "0");
            writer.WriteAttributeString("TextStyle", "0");

            if (effectiveMaster != null) {
                writer.WriteAttributeString("Master", GetPackageMasterId(packageMasters, effectiveMaster));
            }

            KeyValuePair<string, string>? originalIdEntry = GetOriginalIdEntry(persistedIds, shape.Id);
            if (effectiveMaster != null && !isGroup) {
                WriteMasterBackedShapeBody(writer, ns, shape, effectiveMaster, originalIdEntry);
            } else {
                WriteStandaloneShapeBody(writer, ns, shape, isGroup, originalIdEntry);
            }

            if (isGroup && shape.Children.Count > 0) {
                writer.WriteStartElement("Shapes", ns);
                foreach (VisioShape child in shape.Children) {
                    WriteShapeElement(writer, ns, child, persistedIds, effectiveMasters, packageMasters);
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private void WriteMasterBackedShapeBody(XmlWriter writer, string ns, VisioShape shape, VisioMaster master, KeyValuePair<string, string>? originalIdEntry) {
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
                WriteCell(writer, ns, "LineWeight", shape.LineWeight);
                WriteCell(writer, ns, "LinePattern", shape.LinePattern);
                WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
                WriteCell(writer, ns, "FillPattern", shape.FillPattern);
                WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
                WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                WriteDataSection(writer, ns, shape.Data, originalIdEntry);
                WriteTextElement(writer, ns, shape.Text);
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
            WriteConnectionSection(writer, ns, shape.ConnectionPoints);
            WriteDataSection(writer, ns, shape.Data, originalIdEntry);
            WriteTextElement(writer, ns, shape.Text);
        }

        private static void WriteStandaloneShapeBody(XmlWriter writer, string ns, VisioShape shape, bool isGroup, KeyValuePair<string, string>? originalIdEntry) {
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
                 WriteMasterGeometry(writer, ns, shape.NameU, width, height);
              }
            WriteConnectionSection(writer, ns, shape.ConnectionPoints);
            WriteDataSection(writer, ns, shape.Data, originalIdEntry);
            WriteTextElement(writer, ns, shape.Text);
        }

        private static Dictionary<string, string> BuildPersistedIdMap(VisioPage page) {
            Dictionary<string, string> map = new(StringComparer.Ordinal);
            HashSet<int> usedIds = new();

            void Reserve(string originalId) {
                if (map.ContainsKey(originalId)) {
                    return;
                }

                if (int.TryParse(originalId, out int numericId) && numericId >= 0 && usedIds.Add(numericId)) {
                    map[originalId] = originalId;
                    return;
                }

                int nextId = 1;
                while (usedIds.Contains(nextId)) {
                    nextId++;
                }
                usedIds.Add(nextId);
                map[originalId] = nextId.ToString(CultureInfo.InvariantCulture);
            }

            void VisitShape(VisioShape shape) {
                Reserve(shape.Id);
                foreach (VisioShape child in shape.Children) {
                    VisitShape(child);
                }
            }

            foreach (VisioShape shape in page.Shapes) {
                VisitShape(shape);
            }
            foreach (VisioConnector connector in page.Connectors) {
                Reserve(connector.Id);
            }

            return map;
        }

        private static string GetPersistedId(IReadOnlyDictionary<string, string> persistedIds, string originalId) {
            return persistedIds.TryGetValue(originalId, out string? persistedId) ? persistedId : originalId;
        }

        private static KeyValuePair<string, string>? GetOriginalIdEntry(IReadOnlyDictionary<string, string> persistedIds, string originalId) {
            string persistedId = GetPersistedId(persistedIds, originalId);
            return string.Equals(persistedId, originalId, StringComparison.Ordinal)
                ? null
                : new KeyValuePair<string, string>(OriginalIdPropName, originalId);
        }

        private static void WriteConnectorGeometry(XmlWriter writer, string ns, VisioConnector connector, double startX, double startY, double endX, double endY) {
            if (connector.Kind == ConnectorKind.Dynamic) {
                return;
            }

            if (connector.PreservedGeometrySections.Count > 0) {
                foreach (XElement section in connector.PreservedGeometrySections) {
                    XElement clone = new(section);
                    using var reader = clone.CreateReader();
                    writer.WriteNode(reader, false);
                }
                return;
            }

            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", startX);
            WriteCell(writer, ns, "Y", startY);
            writer.WriteEndElement();

            switch (connector.Kind) {
                case ConnectorKind.RightAngle:
                    writer.WriteStartElement("Row", ns);
                    writer.WriteAttributeString("T", "LineTo");
                    WriteCell(writer, ns, "X", startX);
                    WriteCell(writer, ns, "Y", endY);
                    writer.WriteEndElement();

                    writer.WriteStartElement("Row", ns);
                    writer.WriteAttributeString("T", "LineTo");
                    WriteCell(writer, ns, "X", endX);
                    WriteCell(writer, ns, "Y", endY);
                    writer.WriteEndElement();
                    break;
                case ConnectorKind.Curved:
                case ConnectorKind.Straight:
                default:
                    writer.WriteStartElement("Row", ns);
                    writer.WriteAttributeString("T", "LineTo");
                    WriteCell(writer, ns, "X", endX);
                    WriteCell(writer, ns, "Y", endY);
                    writer.WriteEndElement();
                    break;
            }

            writer.WriteEndElement();
        }

        private static void ValidatePagesForSave(IEnumerable<VisioPage> pages) {
            foreach (VisioPage page in pages) {
                HashSet<string> ids = new(StringComparer.Ordinal);

                void Reserve(string id, string kind) {
                    if (string.IsNullOrWhiteSpace(id)) {
                        throw new InvalidOperationException($"{kind} id cannot be null or whitespace on page '{page.Name}'.");
                    }

                    if (!ids.Add(id)) {
                        throw new InvalidOperationException($"Duplicate {kind.ToLowerInvariant()} id '{id}' found on page '{page.Name}'.");
                    }
                }

                void VisitShape(VisioShape shape) {
                    Reserve(shape.Id, "Shape");
                    foreach (VisioShape child in shape.Children) {
                        VisitShape(child);
                    }
                }

                foreach (VisioShape shape in page.Shapes) {
                    VisitShape(shape);
                }

                foreach (VisioConnector connector in page.Connectors) {
                    Reserve(connector.Id, "Connector");
                }
            }
        }

        private Dictionary<string, VisioMaster> BuildEffectiveShapeMasterMap(VisioPage page) {
            Dictionary<string, VisioMaster> map = new(StringComparer.Ordinal);

            void Visit(VisioShape shape) {
                VisioMaster? effectiveMaster = ResolveEffectiveMaster(shape);
                if (effectiveMaster != null) {
                    map[shape.Id] = effectiveMaster;
                }

                foreach (VisioShape child in shape.Children) {
                    Visit(child);
                }
            }

            foreach (VisioShape shape in page.Shapes) {
                Visit(shape);
            }

            return map;
        }

        private static void AddMastersInShapeOrder(IEnumerable<VisioShape> shapes, IReadOnlyDictionary<string, VisioMaster> pageMasters, IList<VisioMaster> masterCandidates) {
            foreach (VisioShape shape in shapes) {
                if (pageMasters.TryGetValue(shape.Id, out VisioMaster? master)) {
                    masterCandidates.Add(master);
                }

                if (shape.Children.Count > 0) {
                    AddMastersInShapeOrder(shape.Children, pageMasters, masterCandidates);
                }
            }
        }

        private VisioMaster? ResolveEffectiveMaster(VisioShape shape) {
            if (shape.Master != null) {
                return shape.Master;
            }

            if (!UseMastersByDefault) {
                return null;
            }

            string shapeNameU = shape.NameU?.Trim() ?? string.Empty;
            if (shapeNameU.Length == 0) {
                return null;
            }

            return TryEnsureBuiltinMaster(shapeNameU, out VisioMaster? master) ? master : null;
        }

        private static List<PackageMasterEntry> CreatePackageMasterEntries(IEnumerable<VisioMaster> masters) {
            List<PackageMasterEntry> entries = new();
            HashSet<int> usedIds = new();

            foreach (VisioMaster master in masters) {
                if (entries.Any(entry => ReferenceEquals(entry.Master, master))) {
                    continue;
                }

                int packageId;
                if (int.TryParse(master.Id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int suggestedId) &&
                    suggestedId > 0 &&
                    usedIds.Add(suggestedId)) {
                    packageId = suggestedId;
                } else {
                    packageId = usedIds.Count == 0 ? 1 : usedIds.Max() + 1;
                    while (!usedIds.Add(packageId)) {
                        packageId++;
                    }
                }

                entries.Add(new PackageMasterEntry {
                    Master = master,
                    PackageId = packageId.ToString(CultureInfo.InvariantCulture),
                    PartNumber = entries.Count + 1
                });
            }

            return entries;
        }

        private static string GetPackageMasterId(IReadOnlyList<PackageMasterEntry> packageMasters, VisioMaster master) {
            foreach (PackageMasterEntry entry in packageMasters) {
                if (ReferenceEquals(entry.Master, master)) {
                    return entry.PackageId;
                }
            }

            throw new InvalidOperationException($"Master '{master.NameU}' is not registered in the package.");
        }

        private static VisioMaster? TryGetEffectiveMaster(IReadOnlyDictionary<string, VisioMaster> effectiveMasters, VisioShape shape) {
            return effectiveMasters.TryGetValue(shape.Id, out VisioMaster? master) ? master : null;
        }

        private static void WriteMasterGeometry(XmlWriter writer, string ns, string? masterNameU, double width, double height) {
            if (TryGetBuiltinMasterDefinition(masterNameU, out BuiltinMasterDefinition? definition) && definition != null) {
                 switch (definition.GeometryKind) {
                     case BuiltinGeometryKind.Ellipse:
                         WriteEllipseGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Diamond:
                         WriteDiamondGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Triangle:
                         WriteTriangleGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Parallelogram:
                         WriteParallelogramGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Hexagon:
                         WriteHexagonGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Trapezoid:
                         WriteTrapezoidGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.OffPageReference:
                         WriteOffPageReferenceGeometry(writer, ns, width, height);
                         return;
                 }
             }

            WriteRectangleGeometry(writer, ns, width, height);
        }

        private static void WriteMasterPageSheet(XmlWriter writer, string ns, BuiltinMasterDefinition? definition) {
            writer.WriteStartElement("PageSheet", ns);
            writer.WriteAttributeString("LineStyle", "0");
            writer.WriteAttributeString("FillStyle", "0");
            writer.WriteAttributeString("TextStyle", "0");
            WritePageCell(writer, ns, "PageWidth", 3.937007874015748, "MM");
            WritePageCell(writer, ns, "PageHeight", 3.937007874015748, "MM");
            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724);
            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724);
            WritePageCell(writer, ns, "PageScale", 0.03937007874015748, "MM");
            WritePageCell(writer, ns, "DrawingScale", 0.03937007874015748, "MM");
            WritePageCell(writer, ns, "DrawingSizeType", definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector ? 4 : 0);
            WritePageCell(writer, ns, "DrawingScaleType", 0);
            WritePageCell(writer, ns, "InhibitSnap", 0);
            WritePageCell(writer, ns, "PageLockReplace", 0, "BOOL");
            WritePageCell(writer, ns, "PageLockDuplicate", 0, "BOOL");
            WritePageCell(writer, ns, "UIVisibility", 0);
            WritePageCell(writer, ns, "ShdwType", 0);
            WritePageCell(writer, ns, "ShdwObliqueAngle", 0);
            WritePageCell(writer, ns, "ShdwScaleFactor", 1);
            WritePageCell(writer, ns, "DrawingResizeType", definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector ? 0 : 1);
            string? shapeKeywords = definition?.ShapeKeywords;
            if (!string.IsNullOrWhiteSpace(shapeKeywords)) {
                WriteStringCell(writer, ns, "ShapeKeywords", shapeKeywords!);
            }
            if (definition?.AddConnectorLayer == true) {
                writer.WriteStartElement("Section", ns);
                writer.WriteAttributeString("N", "Layer");
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("IX", "0");
                WriteStringCell(writer, ns, "Name", "Connector");
                WriteCell(writer, ns, "Color", 255);
                WriteCell(writer, ns, "Status", 0);
                WriteCell(writer, ns, "Visible", 1);
                WriteCell(writer, ns, "Print", 1);
                WriteCell(writer, ns, "Active", 0);
                WriteCell(writer, ns, "Lock", 0);
                WriteCell(writer, ns, "Snap", 1);
                WriteCell(writer, ns, "Glue", 1);
                WriteStringCell(writer, ns, "NameUniv", "Connector");
                WriteCell(writer, ns, "ColorTrans", 0);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteMasterUserSection(XmlWriter writer, string ns) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "User");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("N", "visVersion");
            WriteCell(writer, ns, "Value", 15);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteMasterCharacterSection(XmlWriter writer, string ns) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Character");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("IX", "0");
            WriteCell(writer, ns, "Size", 0.1388888888888889, "PT", null);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteDefaultTextBlock(XmlWriter writer, string ns, double width, double height) {
            WriteCell(writer, ns, "TxtPinX", width / 2.0, "MM", "Width*0.5");
            WriteCell(writer, ns, "TxtPinY", height / 2.0, "MM", "Height*0.5");
            WriteCell(writer, ns, "TxtWidth", width * 0.875, "MM", "Width*0.875");
            WriteCell(writer, ns, "TxtHeight", height * 0.75, "MM", "Height*0.75");
            WriteCell(writer, ns, "TxtLocPinX", width * 0.4375, "MM", "TxtWidth*0.5");
            WriteCell(writer, ns, "TxtLocPinY", height * 0.375, "MM", "TxtHeight*0.5");
            WriteCell(writer, ns, "TxtAngle", 0);
        }

        private static void WriteConnectorControlSection(XmlWriter writer, string ns, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Control");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("N", "TextPosition");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", -height, null, "Controls.TextPosition.Y");
            WriteCell(writer, ns, "XDyn", 0, null, "Controls.TextPosition");
            WriteCell(writer, ns, "YDyn", -height, null, "Controls.TextPosition.Y");
            WriteCell(writer, ns, "XCon", 5);
            WriteCell(writer, ns, "YCon", 0);
            WriteCell(writer, ns, "CanGlue", 0);
            WriteStringCell(writer, ns, "Prompt", "Reposition Text");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private sealed class PackageMasterEntry {
            public VisioMaster Master { get; set; } = null!;
            public string PackageId { get; set; } = string.Empty;
            public int PartNumber { get; set; }
        }
    }
}
