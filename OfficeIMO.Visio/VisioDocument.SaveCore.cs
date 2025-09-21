using System;
using System.Collections.Generic;
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
            int masterCount;
            using (Package package = Package.Open(filePath, FileMode.Create)) {
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

                List<VisioMaster> masters = pagesToSave.SelectMany(p => p.Shapes)
                    .Where(s => s.Master != null)
                    .Select(s => s.Master!)
                    .GroupBy(m => m.Id)
                    .Select(g => g.First())
                    .ToList();

                PackagePart? mastersPart = null;
                if (masters.Count > 0) {
                    Uri mastersUri = new("/visio/masters/masters.xml", UriKind.Relative);
                    mastersPart = package.CreatePart(mastersUri, "application/vnd.ms-visio.masters+xml");
                    documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative), TargetMode.Internal, MastersRelationshipType, "rId4");

                    for (int i = 0; i < masters.Count; i++) {
                        VisioMaster master = masters[i];
                        Uri masterUri = new($"/visio/masters/master{i + 1}.xml", UriKind.Relative);
                        PackagePart masterPart = package.CreatePart(masterUri, "application/vnd.ms-visio.master+xml");
                        mastersPart.CreateRelationship(new Uri($"master{i + 1}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{i + 1}");
                        foreach ((_, PackagePart part, _) in pageParts) {
                            part.CreateRelationship(new Uri($"../masters/master{i + 1}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{i + 1}");
                        }

                        using (XmlWriter writer = XmlWriter.Create(masterPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("MasterContents", ns);
                            writer.WriteStartElement("Shapes", ns);
                            VisioShape s = master.Shape;
                            double masterWidth = s.Width > 0 ? s.Width : 1;
                            double masterHeight = s.Height > 0 ? s.Height : 1;
                            s.Width = masterWidth;
                            s.Height = masterHeight;
                            if (Math.Abs(s.LocPinX) < double.Epsilon) {
                                s.LocPinX = masterWidth / 2;
                            }
                            if (Math.Abs(s.LocPinY) < double.Epsilon) {
                                s.LocPinY = masterHeight / 2;
                            }
                            writer.WriteStartElement("Shape", ns);
                            writer.WriteAttributeString("ID", "1");
                            string masterShapeName = s.Name ?? s.NameU ?? "MasterShape";
                            writer.WriteAttributeString("Name", masterShapeName);
                            writer.WriteAttributeString("NameU", master.NameU);
                            writer.WriteAttributeString("Type", "Shape");
                            writer.WriteAttributeString("LineStyle", "0");
                            writer.WriteAttributeString("FillStyle", "0");
                            writer.WriteAttributeString("TextStyle", "0");
                            WriteXForm(writer, ns, s, masterWidth, masterHeight);
                            // Always specify line weight so that shapes are visible
                            WriteCell(writer, ns, "LineWeight", s.LineWeight);
                            WriteCell(writer, ns, "LinePattern", s.LinePattern);
                            WriteCellValue(writer, ns, "LineColor", s.LineColor.ToVisioHex());
                            WriteCell(writer, ns, "FillPattern", s.FillPattern);
                            WriteCellValue(writer, ns, "FillForegnd", s.FillColor.ToVisioHex());
                            WriteRectangleGeometry(writer, ns, masterWidth, masterHeight);
                            WriteConnectionSection(writer, ns, s.ConnectionPoints);
                            WriteDataSection(writer, ns, s.Data);
                            WriteTextElement(writer, ns, s.Text);
                            writer.WriteEndElement();
                            writer.WriteEndElement(); // Shapes
                            writer.WriteEndElement(); // MasterContents
                            writer.WriteEndDocument();
                        }
                    }

                    // Write masters list (masters.xml)
                    using (XmlWriter writer = XmlWriter.Create(mastersPart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Masters", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        for (int i = 0; i < masters.Count; i++) {
                            VisioMaster m = masters[i];
                            writer.WriteStartElement("Master", ns);
                            writer.WriteAttributeString("ID", m.Id);
                            writer.WriteAttributeString("NameU", m.NameU);
                            writer.WriteStartElement("Rel", ns);
                            writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{i + 1}");
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
                        writer.WriteAttributeString("ViewScale", XmlConvert.ToString(page.ViewScale));
                        writer.WriteAttributeString("ViewCenterX", XmlConvert.ToString(page.ViewCenterX));
                        writer.WriteAttributeString("ViewCenterY", XmlConvert.ToString(page.ViewCenterY));

                        writer.WriteStartElement("PageSheet", ns);
                        writer.WriteAttributeString("LineStyle", "0");
                        writer.WriteAttributeString("FillStyle", "0");
                        writer.WriteAttributeString("TextStyle", "0");

                        bool useUnits = page.Width != 8.26771653543307 || page.Height != 11.69291338582677;
                        if (useUnits) {
                            // Match asset semantics: write inch values but mark as MM for page size and shadow offsets
                            WritePageCell(writer, ns, "PageWidth", page.Width, "MM");
                            WritePageCell(writer, ns, "PageHeight", page.Height, "MM");
                            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724, "MM");
                            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724, "MM");
                        } else {
                            WritePageCell(writer, ns, "PageWidth", page.Width);
                            WritePageCell(writer, ns, "PageHeight", page.Height);
                            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724);
                            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724);
                        }
                        WritePageCell(writer, ns, "PageScale", 0.03937007874015748, "MM");
                        WritePageCell(writer, ns, "DrawingScale", 0.03937007874015748, "MM");
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
                        writer.WriteStartDocument();
                        writer.WriteStartElement("PageContents", ns);
                        writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        writer.WriteAttributeString("xml", "space", null, "preserve");
                        if (page.Shapes.Count > 0 || page.Connectors.Count > 0) {
                            writer.WriteStartElement("Shapes", ns);

                            foreach (VisioShape shape in page.Shapes) {
                                writer.WriteStartElement("Shape", ns);
                                writer.WriteAttributeString("ID", shape.Id);
                                string shapeName = shape.Name ?? shape.NameU ?? $"Shape{shape.Id}";
                                writer.WriteAttributeString("Name", shapeName);
                                writer.WriteAttributeString("NameU", shape.NameU ?? shape.Master?.NameU ?? shapeName);
                                writer.WriteAttributeString("Type", "Shape");
                                writer.WriteAttributeString("LineStyle", "0");
                                writer.WriteAttributeString("FillStyle", "0");
                                writer.WriteAttributeString("TextStyle", "0");
                                if (shape.Master != null) {
                                    writer.WriteAttributeString("Master", shape.Master.Id);
                                    double width = shape.Width;
                                    if (width <= 0 && shape.Master.Shape.Width > 0) {
                                        width = shape.Master.Shape.Width;
                                    }
                                    if (width <= 0) {
                                        width = 1;
                                    }
                                    double height = shape.Height;
                                    if (height <= 0 && shape.Master.Shape.Height > 0) {
                                        height = shape.Master.Shape.Height;
                                    }
                                    if (height <= 0) {
                                        height = 1;
                                    }
                                    shape.Width = width;
                                    shape.Height = height;
                                    if (Math.Abs(shape.LocPinX) < double.Epsilon) {
                                        shape.LocPinX = width / 2;
                                    }
                                    if (Math.Abs(shape.LocPinY) < double.Epsilon) {
                                        shape.LocPinY = height / 2;
                                    }
                                    WriteXForm(writer, ns, shape, width, height);
                                    // Always include line weight to avoid invisible shapes
                                    WriteCell(writer, ns, "LineWeight", shape.LineWeight);
                                    WriteCell(writer, ns, "LinePattern", shape.LinePattern);
                                    WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
                                    WriteCell(writer, ns, "FillPattern", shape.FillPattern);
                                    WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
                                    WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                                    WriteDataSection(writer, ns, shape.Data);
                                    WriteTextElement(writer, ns, shape.Text);
                                } else {
                                    double width = shape.Width > 0 ? shape.Width : 1;
                                    double height = shape.Height > 0 ? shape.Height : 1;
                                    shape.Width = width;
                                    shape.Height = height;
                                    if (Math.Abs(shape.LocPinX) < double.Epsilon) {
                                        shape.LocPinX = width / 2;
                                    }
                                    if (Math.Abs(shape.LocPinY) < double.Epsilon) {
                                        shape.LocPinY = height / 2;
                                    }
                                    WriteXForm(writer, ns, shape, width, height);
                                    // Always include line weight to avoid invisible shapes
                                    WriteCell(writer, ns, "LineWeight", shape.LineWeight);
                                    WriteCell(writer, ns, "LinePattern", shape.LinePattern);
                                    WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
                                    WriteCell(writer, ns, "FillPattern", shape.FillPattern);
                                    WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
                                    string? kind = shape.NameU?.Trim();
                                    if (string.Equals(kind, "Ellipse", StringComparison.OrdinalIgnoreCase) || string.Equals(kind, "Circle", StringComparison.OrdinalIgnoreCase)) {
                                        WriteEllipseGeometry(writer, ns, width, height);
                                    } else if (string.Equals(kind, "Diamond", StringComparison.OrdinalIgnoreCase)) {
                                        WriteDiamondGeometry(writer, ns, width, height);
                                    } else if (string.Equals(kind, "Triangle", StringComparison.OrdinalIgnoreCase)) {
                                        WriteTriangleGeometry(writer, ns, width, height);
                                    } else {
                                        WriteRectangleGeometry(writer, ns, width, height);
                                    }
                                    WriteConnectionSection(writer, ns, shape.ConnectionPoints);
                                    WriteDataSection(writer, ns, shape.Data);
                                    WriteTextElement(writer, ns, shape.Text);
                                }
                                writer.WriteEndElement();
                            }

                            foreach (VisioConnector connector in page.Connectors) {
                                writer.WriteStartElement("Shape", ns);
                                writer.WriteAttributeString("ID", connector.Id);
                                // Mark as connector so the loader can distinguish it.
                                writer.WriteAttributeString("Name", "Connector");
                                writer.WriteAttributeString("NameU", "Connector");
                                writer.WriteAttributeString("LineStyle", "0");
                                writer.WriteAttributeString("FillStyle", "0");
                                writer.WriteAttributeString("TextStyle", "0");
                                (double startX, double startY) = connector.FromConnectionPoint != null
                                    ? connector.From.GetAbsolutePoint(connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y)
                                    : (connector.From.PinX + connector.From.Width / 2.0, connector.From.PinY);
                                (double endX, double endY) = connector.ToConnectionPoint != null
                                    ? connector.To.GetAbsolutePoint(connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y)
                                    : (connector.To.PinX - connector.To.Width / 2.0, connector.To.PinY);
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

                                if (connector.Kind != ConnectorKind.Dynamic) {
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

                                WriteTextElement(writer, ns, connector.Label);
                                writer.WriteEndElement();
                            }

                            writer.WriteEndElement(); // Shapes

                            if (page.Connectors.Count > 0) {
                                writer.WriteStartElement("Connects", ns);
                                foreach (VisioConnector connector in page.Connectors) {
                                    writer.WriteStartElement("Connect", ns);
                                    writer.WriteAttributeString("FromSheet", connector.Id);
                                    writer.WriteAttributeString("FromCell", "BeginX");
                                    writer.WriteAttributeString("ToSheet", connector.From.Id);
                                    writer.WriteAttributeString("ToCell", GetConnectionCell(connector.From, connector.FromConnectionPoint));
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("Connect", ns);
                                    writer.WriteAttributeString("FromSheet", connector.Id);
                                    writer.WriteAttributeString("FromCell", "EndX");
                                    writer.WriteAttributeString("ToSheet", connector.To.Id);
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
            }

            FixContentTypes(filePath, masterCount, includeTheme, pageCount);
        }
    }
}
