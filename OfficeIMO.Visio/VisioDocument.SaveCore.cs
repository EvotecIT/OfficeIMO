using OfficeIMO.Core.Internal;
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
    // Save core implementation for VisioDocument.
    public partial class VisioDocument {
        /// <summary>
        /// Core save routine that writes the VSdx structure.
        /// </summary>
        /// <param name="filePath">Target path.</param>
        private void SaveInternalCore(string filePath) {
            OfficeFileCommit.WriteAllBytes(filePath, CreatePackageBytes());
        }

        /// <summary>
        /// Core save routine that writes the VSDX structure to a stream.
        /// </summary>
        /// <param name="destination">Target stream.</param>
        private void SaveInternalCore(Stream destination) {
            using MemoryStream packageStream = CreatePackageStream();
            OfficeStreamWriter.Write(destination, output => packageStream.CopyTo(output));
        }

        private byte[] CreatePackageBytes() {
            using MemoryStream packageStream = CreatePackageStream();
            return packageStream.ToArray();
        }

        private MemoryStream CreatePackageStream() {
            ApplySignatureMutationPolicy();
            bool includeTheme = PackageTheme != null;
            List<VisioPage> pagesToSave = _pages.Count > 0
                ? _pages
                : IsStencil
                    ? new List<VisioPage>()
                    : new List<VisioPage> { new VisioPage("Page-1") { Id = 0 } };
            bool includeComments = pagesToSave.Any(page => page.Comments.Count > 0);
            PrepareTextFontFaceNames(pagesToSave);
            ValidatePagesForSave(pagesToSave);
            if (includeComments) {
                ValidateCommentsForSave(pagesToSave, BuildEffectivePageMastersForSave(pagesToSave));
            }

            int pageCount = pagesToSave.Count;
            List<string> pagePartNames = new();
            int masterCount;

            var packageStream = new MemoryStream();
            try {
                using (Package package = Package.Open(packageStream, FileMode.Create, FileAccess.ReadWrite)) {
                    masterCount = WritePackage(package, includeTheme, includeComments, pagesToSave, pageCount, pagePartNames);
                }

                FixContentTypes(packageStream, masterCount, includeTheme,
                    includeComments, pagePartNames, _packageType,
                    _vbaProjectBytes != null && _vbaProjectBytes.Length > 0,
                    _vbaProjectContentType, _vbaProjectPartUri,
                    _preservedVbaParts.Values);

                packageStream.Seek(0, SeekOrigin.Begin);
                return packageStream;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

        private int WritePackage(
            Package package,
            bool includeTheme,
            bool includeComments,
            List<VisioPage> pagesToSave,
            int pageCount,
            List<string> pagePartNames) {
            int masterCount;
            Uri documentUri = new("/visio/document.xml", UriKind.Relative);
            PackagePart documentPart = package.CreatePart(documentUri,
                VisioPackageFormat.GetContentType(_packageType));
            // Package-relative targets let independent readers locate the document stream directly.
            package.CreateRelationship(new Uri(documentUri.OriginalString.TrimStart('/'), UriKind.Relative),
                TargetMode.Internal, DocumentRelationshipType, "rId1");

                Uri coreUri = new("/docProps/core.xml", UriKind.Relative);
                PackagePart corePart = package.CreatePart(coreUri, "application/vnd.openxmlformats-package.core-properties+xml");
                package.CreateRelationship(coreUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "rId2");

                Uri appUri = new("/docProps/app.xml", UriKind.Relative);
                PackagePart appPart = package.CreatePart(appUri, "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                package.CreateRelationship(appUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "rId3");

                Uri customUri = new("/docProps/custom.xml", UriKind.Relative);
                PackagePart customPart = package.CreatePart(customUri, "application/vnd.openxmlformats-officedocument.custom-properties+xml");
                package.CreateRelationship(customUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "rId4");

                Uri pagesUri = new("/visio/pages/pages.xml", UriKind.Relative);
                PackagePart pagesPart = package.CreatePart(pagesUri, PagesContentType);
                documentPart.CreateRelationship(new Uri("pages/pages.xml", UriKind.Relative), TargetMode.Internal, PagesRelationshipType, "rId1");

                Uri windowsUri = new("/visio/windows.xml", UriKind.Relative);
                PackagePart windowsPart = package.CreatePart(windowsUri, WindowsContentType);
                documentPart.CreateRelationship(new Uri("windows.xml", UriKind.Relative), TargetMode.Internal, WindowsRelationshipType, "rId2");

                int nextDocumentRelationshipId = 3;
                if (_vbaProjectBytes != null && _vbaProjectBytes.Length > 0) {
                    if (!VisioPackageFormat.IsMacroEnabled(_packageType)) {
                        throw new InvalidOperationException(
                            "A preserved VBA project requires a macro-enabled Visio package type.");
                    }
                    WriteVbaSubtree(package);
                    documentPart.CreateRelationship(
                        PackUriHelper.GetRelativeUri(documentPart.Uri,
                            _vbaProjectPartUri),
                        TargetMode.Internal, VbaProjectRelationshipType,
                        $"rId{nextDocumentRelationshipId++}");
                }
                PackagePart? themePart = null;
                if (includeTheme) {
                    Uri themeUri = new("/visio/theme/theme1.xml", UriKind.Relative);
                    themePart = package.CreatePart(themeUri, ThemeContentType);
                    documentPart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal, ThemeRelationshipType, $"rId{nextDocumentRelationshipId++}");
                }

                PackagePart? commentsPart = null;
                if (includeComments) {
                    Uri commentsUri = new("/visio/comments.xml", UriKind.Relative);
                    commentsPart = package.CreatePart(commentsUri, CommentsContentType);
                    documentPart.CreateRelationship(new Uri("comments.xml", UriKind.Relative), TargetMode.Internal, CommentsRelationshipType, $"rId{nextDocumentRelationshipId++}");
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

                if (themePart != null && PackageTheme != null) {
                    if (PackageTheme.TemplateXml != null) {
                        XDocument themeXml = new(PackageTheme.TemplateXml);
                        if (themeXml.Root != null) {
                            themeXml.Root.SetAttributeValue("name", PackageTheme.Name);
                        }

                        using Stream s = themePart.GetStream(FileMode.Create, FileAccess.Write);
                        using StreamWriter sw = new(s, new UTF8Encoding(false));
                        sw.Write(themeXml.Declaration + Environment.NewLine + themeXml.ToString(SaveOptions.DisableFormatting));
                    } else {
                        using (XmlWriter writer = XmlWriter.Create(themePart.GetStream(FileMode.Create, FileAccess.Write), settings)) {
                            writer.WriteStartDocument();
                            writer.WriteStartElement("a", "theme", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            if (!string.IsNullOrEmpty(PackageTheme.Name)) {
                                writer.WriteAttributeString("name", PackageTheme.Name);
                            }
                            writer.WriteEndElement();
                            writer.WriteEndDocument();
                        }
                    }
                }

                // Complete document XML after shape parts assign their native identities.
                XDocument documentXml = CreateVisioDocumentXml(
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

                WritePackageMetadata(appPart, customPart, windowsPart, pagesToSave);

                Dictionary<VisioPage, Dictionary<string, VisioMaster>> effectivePageMasters = new();
                List<VisioMaster> masterCandidates = new();
                if (IsStencil || IsTemplate) masterCandidates.AddRange(_registeredMasters);
                foreach (VisioPage page in pagesToSave) {
                    Dictionary<string, VisioMaster> pageMasters = BuildEffectiveShapeMasterMap(page);
                    effectivePageMasters[page] = pageMasters;
                    AddMastersInShapeOrder(page.Shapes, pageMasters, masterCandidates);

                    foreach (VisioConnector connector in page.Connectors) {
                        VisioMaster? master = ResolveEffectiveMaster(connector);
                        if (master != null) masterCandidates.Add(master);
                    }
                }

                List<PackageMasterEntry> masters = CreatePackageMasterEntries(masterCandidates);

                PackagePart? mastersPart = null;
                if (masters.Count > 0) {
                    Uri mastersUri = new("/visio/masters/masters.xml", UriKind.Relative);
                    mastersPart = package.CreatePart(mastersUri, "application/vnd.ms-visio.masters+xml");
                    documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative), TargetMode.Internal, MastersRelationshipType, $"rId{nextDocumentRelationshipId++}");

                    for (int i = 0; i < masters.Count; i++) {
                        PackageMasterEntry entry = masters[i];
                        VisioMaster master = entry.Master;
                        Uri masterUri = new($"/visio/masters/master{entry.PartNumber}.xml", UriKind.Relative);
                        PackagePart masterPart = package.CreatePart(masterUri, "application/vnd.ms-visio.master+xml");
                        mastersPart.CreateRelationship(new Uri($"master{entry.PartNumber}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{entry.PartNumber}");
                        foreach ((_, PackagePart part, _) in pageParts) {
                            part.CreateRelationship(new Uri($"../masters/master{entry.PartNumber}.xml", UriKind.Relative), TargetMode.Internal, MasterRelationshipType, $"rId{entry.PartNumber}");
                        }

                        var foreignResources = master.ForeignResources.Concat(ShapeForeignResources(new[] { master.Shape })).ToList();
                        if ((foreignResources.Count > 0 && master.LoadedModelShapeXml == null) ||
                            (master.RawMasterContentXml == null && RequiresCompleteMasterShape(master.Shape))) {
                            WriteModeledMasterContent(masterPart, master, ns, settings);
                            WriteForeignResources(masterPart, foreignResources);
                            continue;
                        }
                        if (master.RawMasterContentXml != null) {
                            WriteLoadedMasterContent(masterPart, master);
                            WriteRawMasterRelationships(package, masterPart, master, entry.PartNumber);
                            WriteForeignResources(masterPart, foreignResources);
                            continue;
                        }

                        WriteSimpleMasterContent(masterPart, master, ns, settings);
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
                            WriteMasterCatalogAttributes(writer, m, entry.PackageId, masterDefinition);
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

                if (commentsPart != null) {
                    WriteCommentsPart(commentsPart, pagesToSave, effectivePageMasters);
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

                        WritePageSheet(writer, ns, page);
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
                    WriteForeignResources(pagePart, page.ForeignResources.Concat(ShapeForeignResources(page.Shapes)).Concat(page.Connectors.SelectMany(connector => connector.ForeignResources)));
                }
                CollectNativeCellMetadata(documentXml, package, pageParts, masters, pagesPart, mastersPart);
                WriteNativeFontNames(documentXml, package, pageParts, masters, pagesPart, mastersPart);
                using (Stream stream = documentPart.GetStream(FileMode.Create, FileAccess.Write))
                using (StreamWriter writer = new(stream, new UTF8Encoding(false)))
                    writer.Write(documentXml.Declaration + Environment.NewLine + documentXml.ToString(SaveOptions.DisableFormatting));
                masterCount = masters.Count;
                pagePartNames.Clear();
                pagePartNames.AddRange(pageParts
                    .Select(part => part.Part.Uri.OriginalString)
                    .Distinct(StringComparer.OrdinalIgnoreCase));

            return masterCount;
        }

        private void WriteVbaSubtree(Package package) {
            IEnumerable<PreservedVbaPart> sourceParts =
                _preservedVbaParts.Count > 0
                    ? _preservedVbaParts.Values
                    : new[] {
                        new PreservedVbaPart(_vbaProjectPartUri,
                            string.IsNullOrWhiteSpace(_vbaProjectContentType)
                                ? VbaProjectContentType
                                : _vbaProjectContentType!,
                            _vbaProjectBytes!,
                            Array.Empty<PreservedVbaRelationship>())
                    };
            PreservedVbaPart[] parts = sourceParts.ToArray();
            foreach (PreservedVbaPart source in parts) {
                if (package.PartExists(source.Uri)) {
                    throw new InvalidOperationException(
                        $"The preserved VBA subtree conflicts with generated package part '{source.Uri}'.");
                }
                PackagePart target = package.CreatePart(source.Uri,
                    source.ContentType);
                using Stream stream = target.GetStream(FileMode.Create,
                    FileAccess.Write);
                stream.Write(source.Data, 0, source.Data.Length);
            }
            foreach (PreservedVbaPart source in parts) {
                PackagePart target = package.GetPart(source.Uri);
                foreach (PreservedVbaRelationship relationship in
                         source.Relationships) {
                    target.CreateRelationship(relationship.TargetUri,
                        relationship.TargetMode, relationship.Type,
                        relationship.Id);
                }
            }
        }

    }
}
