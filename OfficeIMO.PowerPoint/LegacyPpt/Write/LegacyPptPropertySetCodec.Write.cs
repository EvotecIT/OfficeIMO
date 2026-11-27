using System.Globalization;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPropertySetCodec {
        private const string CustomPropertiesFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

        private static byte[] CreateEmptyPropertySet(Guid formatId) {
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            return OfficeOlePropertySetWriter.CreatePropertySet((formatId,
                OfficeOlePropertySetWriter.CreateSection(properties)));
        }

        private static byte[]? CreateSummaryInformation(
            PackageProperties source) {
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(properties, 2, source.Core("title"));
            OfficeOlePropertySetWriter.AddString(properties, 3, source.Core("subject"));
            OfficeOlePropertySetWriter.AddString(properties, 4, source.Core("creator"));
            OfficeOlePropertySetWriter.AddString(properties, 5, source.Core("keywords"));
            OfficeOlePropertySetWriter.AddString(properties, 6,
                source.Core("description"));
            OfficeOlePropertySetWriter.AddString(properties, 8,
                source.Core("lastModifiedBy"));
            OfficeOlePropertySetWriter.AddString(properties, 9,
                source.Core("revision"));
            if (int.TryParse(source.Application("TotalTime"),
                    NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out int totalMinutes) && totalMinutes >= 0) {
                properties.Add(OfficeOleProperty.FileTimeDuration(10,
                    TimeSpan.FromMinutes(totalMinutes)));
            }
            OfficeOlePropertySetWriter.AddFileTime(properties, 11,
                source.CoreDate("lastPrinted"));
            OfficeOlePropertySetWriter.AddFileTime(properties, 12,
                source.CoreDate("created"));
            OfficeOlePropertySetWriter.AddFileTime(properties, 13,
                source.CoreDate("modified"));
            OfficeOlePropertySetWriter.AddString(properties, 18,
                source.Application("Application"));
            return properties.Count == 1 ? null
                : OfficeOlePropertySetWriter.CreatePropertySet((
                    OfficeOlePropertySetWriter.SummaryInformationFormatId,
                    OfficeOlePropertySetWriter.CreateSection(properties)));
        }

        private static bool TryCreateDocumentSummaryInformation(
            PackageProperties source, out byte[]? bytes,
            out string? reason) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            OfficeOlePropertySetWriter.AddString(properties, 2,
                source.Core("category"));
            OfficeOlePropertySetWriter.AddString(properties, 3,
                source.Application("PresentationFormat"));
            AddIntegerText(properties, 7, source.Application("Slides"));
            AddIntegerText(properties, 8, source.Application("Notes"));
            AddIntegerText(properties, 9, source.Application("HiddenSlides"));
            OfficeOlePropertySetWriter.AddString(properties, 14,
                source.Application("Manager"));
            OfficeOlePropertySetWriter.AddString(properties, 15,
                source.Application("Company"));
            if (properties.Count > 1) {
                sections.Add((OfficeOlePropertySetWriter
                    .DocumentSummaryInformationFormatId,
                    OfficeOlePropertySetWriter.CreateSection(properties)));
            }
            if (!TryCreateCustomSection(source.Custom,
                    out byte[]? customSection, out reason)) {
                bytes = null;
                return false;
            }
            if (customSection != null) {
                sections.Add((OfficeOlePropertySetWriter
                    .UserDefinedPropertiesFormatId, customSection));
            }
            bytes = sections.Count == 0 ? null
                : OfficeOlePropertySetWriter.CreatePropertySet(
                    sections.ToArray());
            reason = null;
            return true;
        }

        private static void AddIntegerText(ICollection<OfficeOleProperty> target,
            uint id, string? text) {
            if (int.TryParse(text, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out int value)) {
                target.Add(OfficeOleProperty.Int32(id, value));
            }
        }

        private static bool TryCreateCustomSection(
            OfficeXmlNode customRoot, out byte[]? section,
            out string? reason) {
            section = null;
            reason = null;
            List<OfficeXmlNode> source = customRoot.Children(
                OfficeXmlNamespaces.CustomProperties, "property").ToList();
            if (source.Count == 0) return true;
            var properties = new List<OfficeOleProperty> {
                OfficeOleProperty.Int16(
                    OfficeOlePropertySetWriter.CodePagePropertyId, 1200)
            };
            var names = new Dictionary<uint, string>();
            uint propertyId = 2;
            foreach (OfficeXmlNode item in source) {
                string? name = item.Attribute("name");
                if (string.IsNullOrWhiteSpace(name)
                    || !TryCreateOleCustomProperty(propertyId, item,
                        out OfficeOleProperty property)) {
                    reason = "A custom document property has no name or uses an unsupported Open XML variant type.";
                    return false;
                }
                names.Add(propertyId, name!);
                properties.Add(property);
                propertyId++;
            }
            properties.Insert(1, OfficeOleProperty.Dictionary(0, names));
            section = OfficeOlePropertySetWriter.CreateSection(properties);
            return true;
        }

        private static string? VariantText(OfficeXmlNode property, string variant) {
            OfficeXmlNode value = property.Child(OfficeXmlNamespaces.DocPropsVTypes, variant);
            return value.IsNone ? null : value.InnerText;
        }

        private static bool HasVariant(OfficeXmlNode property, string variant) =>
            !property.Child(OfficeXmlNamespaces.DocPropsVTypes, variant).IsNone;

        private static bool TryCreateOleCustomProperty(uint id,
            OfficeXmlNode source, out OfficeOleProperty property) {
            if (HasVariant(source, "lpwstr")) {
                property = OfficeOleProperty.String(id,
                    VariantText(source, "lpwstr") ?? string.Empty);
            } else if (HasVariant(source, "bool")
                       && TryParseBoolean(VariantText(source, "bool"),
                           out bool boolean)) {
                property = OfficeOleProperty.Boolean(id, boolean);
            } else if (HasVariant(source, "filetime")
                       && DateTime.TryParse(VariantText(source, "filetime"),
                           CultureInfo.InvariantCulture,
                           DateTimeStyles.AssumeUniversal
                           | DateTimeStyles.AdjustToUniversal,
                           out DateTime date)) {
                property = OfficeOleProperty.FileTime(id, date);
            } else if (HasVariant(source, "r8")
                       && OfficeRoundTripText.TryParse(VariantText(source, "r8"), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                property = OfficeOleProperty.Double(id, number);
            } else if (HasVariant(source, "r4")
                       && OfficeRoundTripText.TryParse(VariantText(source, "r4"), NumberStyles.Float, CultureInfo.InvariantCulture, out float single)) {
                property = OfficeOleProperty.Float(id, single);
            } else if (HasVariant(source, "i4")
                       && int.TryParse(VariantText(source, "i4"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out int integer)) {
                property = OfficeOleProperty.Int32(id, integer);
            } else if (HasVariant(source, "i8")
                       && long.TryParse(VariantText(source, "i8"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out long integer64)) {
                property = OfficeOleProperty.Int64(id, integer64);
            } else if (HasVariant(source, "ui1")
                       && byte.TryParse(VariantText(source, "ui1"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out byte uint8)) {
                property = OfficeOleProperty.UInt8(id, uint8);
            } else if (HasVariant(source, "ui2")
                       && ushort.TryParse(VariantText(source, "ui2"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out ushort uint16)) {
                property = OfficeOleProperty.UInt16(id, uint16);
            } else if (HasVariant(source, "ui4")
                       && uint.TryParse(VariantText(source, "ui4"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out uint uint32)) {
                property = OfficeOleProperty.UInt32(id, uint32);
            } else if (HasVariant(source, "ui8")
                       && ulong.TryParse(VariantText(source, "ui8"),
                           NumberStyles.Integer,
                           CultureInfo.InvariantCulture, out ulong uint64)) {
                property = OfficeOleProperty.UInt64(id, uint64);
            } else if (HasVariant(source, "blob")
                       && TryParseBase64(VariantText(source, "blob"),
                           out byte[] blob)) {
                property = OfficeOleProperty.Blob(id, blob);
            } else {
                property = default;
                return false;
            }
            return true;
        }

        private static bool TryParseBoolean(string? text, out bool value) {
            if (bool.TryParse(text, out value)) return true;
            if (text == "1" || text == "-1") {
                value = true;
                return true;
            }
            if (text == "0") {
                value = false;
                return true;
            }
            return false;
        }

        private static bool TryParseBase64(string? text, out byte[] value) {
            try {
                value = string.IsNullOrEmpty(text)
                    ? Array.Empty<byte>()
                    : Convert.FromBase64String(text);
                return true;
            } catch (FormatException) {
                value = Array.Empty<byte>();
                return false;
            }
        }

        /// <summary>One custom document property to write: its name, id and <c>vt:</c> value.</summary>
        private sealed class CustomPropertyValue {
            internal CustomPropertyValue(string name, int propertyId, string variant, string text) {
                Name = name;
                PropertyId = propertyId;
                Variant = variant;
                Text = text;
            }

            internal string Name { get; }
            internal int PropertyId { get; }
            internal string Variant { get; }
            internal string Text { get; }
        }

        /// <summary>The document properties of a package, read from a <see cref="LegacyPptPackageView"/>.</summary>
        private sealed class PackageProperties {
            private readonly OfficeXmlNode _core;
            private readonly OfficeXmlNode _application;

            private PackageProperties(OfficeXmlNode core, OfficeXmlNode application, OfficeXmlNode custom) {
                _core = core;
                _application = application;
                Custom = custom;
            }

            /// <summary>The custom properties root, or none.</summary>
            internal OfficeXmlNode Custom { get; }

            internal static PackageProperties Read(LegacyPptPackageView view) => new(
                view.PackageRelatedPart(LegacyPptPackageView.CorePropertiesRelationshipType).Root,
                view.PackageRelatedPart(LegacyPptPackageView.ExtendedPropertiesRelationshipType).Root,
                view.PackageRelatedPart(LegacyPptPackageView.CustomPropertiesRelationshipType).Root);

            /// <summary>A core property's text, or null when the package does not set it.</summary>
            internal string? Core(string localName) {
                OfficeXmlNode element = _core.Child(CoreNamespace(localName), localName);
                return element.IsNone ? null : element.InnerText;
            }

            internal DateTime? CoreDate(string localName) => ParseCoreDate(Core(localName));

            /// <summary>An extended (application) property's text, or null when the package does not set it.</summary>
            internal string? Application(string localName) {
                OfficeXmlNode element = _application.Child(OfficeXmlNamespaces.ExtendedProperties, localName);
                return element.IsNone ? null : element.InnerText;
            }
        }

        /// <summary>The namespace a core property element is in.</summary>
        private static int CoreNamespace(string localName) => localName switch {
            "title" or "subject" or "creator" or "description" or "identifier" or "language" => OfficeXmlNamespaces.DublinCore,
            "created" or "modified" => OfficeXmlNamespaces.DublinCoreTerms,
            _ => OfficeXmlNamespaces.CoreProperties
        };

        private static readonly string[] MappedCoreProperties = {
            "title", "subject", "creator", "keywords", "description", "lastModifiedBy", "revision", "lastPrinted",
            "created", "modified", "category", "contentStatus", "contentType", "identifier", "language", "version"
        };

        private static readonly string[] MappedApplicationProperties = {
            "Application", "TotalTime", "PresentationFormat", "Slides", "Notes", "HiddenSlides", "Manager", "Company"
        };

        private static DateTime? ParseCoreDate(string? text) =>
            text != null && DateTime.TryParse(text, CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime value)
                ? value
                : null;

        /// <summary>Where the binary property sets are written: the package, or an SDK document.</summary>
        private interface IPropertyTarget {
            void ClearMapped();
            void SetCore(string localName, string? value);
            void SetCoreDate(string localName, DateTime? value);
            void EnsureApplication();
            void SetApplication(string localName, string value);
            void SetCustom(IReadOnlyList<CustomPropertyValue> values);
        }

        /// <summary>Writes document properties into an editable package.</summary>
        private sealed class PackagePropertyWriter : IPropertyTarget {
            private const string XsiUri = "http://www.w3.org/2001/XMLSchema-instance";
            private const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
            private const string CustomPropertiesContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
            private static readonly OfficeXmlChildOrder ApplicationOrder = CreateApplicationOrder();
            private readonly OfficeOpcEditablePackage _package;

            internal PackagePropertyWriter(OfficeOpcEditablePackage package) => _package = package;

            private static OfficeXmlChildOrder CreateApplicationOrder() {
                string[] names = {
                    "Template", "Manager", "Company", "Pages", "Words", "Characters", "PresentationFormat", "Lines",
                    "Paragraphs", "Slides", "Notes", "TotalTime", "HiddenSlides", "MMClips", "ScaleCrop", "HeadingPairs",
                    "TitlesOfParts", "LinksUpToDate", "CharactersWithSpaces", "SharedDoc", "HyperlinkBase", "HLinks",
                    "HyperlinksChanged", "DigSig", "Application", "AppVersion", "DocSecurity"
                };
                int[] namespaces = names.Select(_ => OfficeXmlNamespaces.ExtendedProperties).ToArray();
                return new OfficeXmlChildOrder(namespaces, names);
            }

            private string? RelatedPart(string relationshipType) {
                string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(null);
                if (!_package.PartExists(relationshipsPart)) return null;
                OfficeXmlTree tree = _package.GetRelationships(null);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                     node != OfficeXmlTree.NoNode;
                     node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                    if (type == null || !string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(type), relationshipType, StringComparison.Ordinal)
                        || string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }
                    string? part = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart,
                        tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (part != null && _package.PartExists(part)) return part;
                }
                return null;
            }

            private OfficeXmlTree? Core(bool create) {
                string? part = RelatedPart(LegacyPptPackageView.CorePropertiesRelationshipType);
                if (part != null) return _package.GetXml(part);
                if (!create) return null;
                string name = UniquePartName("/docProps/core.xml");
                OfficeXmlTree core = _package.CreateXmlPart(name, LegacyPptPackageFingerprint.CorePropertiesContentType,
                    OfficeXmlNamespaces.CoreProperties, "coreProperties", 8);
                core.DeclareNamespace(core.Root, "cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                core.DeclareNamespace(core.Root, "dc", "http://purl.org/dc/elements/1.1/");
                core.DeclareNamespace(core.Root, "dcterms", "http://purl.org/dc/terms/");
                core.DeclareNamespace(core.Root, "dcmitype", "http://purl.org/dc/dcmitype/");
                core.DeclareNamespace(core.Root, "xsi", XsiUri);
                _package.AddRelationship(null, LegacyPptPackageView.CorePropertiesRelationshipType, name.Substring(1));
                return core;
            }

            private OfficeXmlTree? Application(bool create) {
                string? part = RelatedPart(LegacyPptPackageView.ExtendedPropertiesRelationshipType);
                if (part != null) return _package.GetXml(part);
                if (!create) return null;
                string name = UniquePartName("/docProps/app.xml");
                OfficeXmlTree application = _package.CreateXmlPart(name, ExtendedPropertiesContentType,
                    OfficeXmlNamespaces.ExtendedProperties, "Properties", 4);
                application.DeclareNamespace(application.Root, string.Empty, "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
                application.DeclareNamespace(application.Root, "vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                _package.AddRelationship(null, LegacyPptPackageView.ExtendedPropertiesRelationshipType, name.Substring(1));
                return application;
            }

            private string UniquePartName(string requested) {
                if (!_package.PartExists(requested)) return requested;
                string stem = requested.Substring(0, requested.Length - ".xml".Length);
                for (int index = 2; ; index++) {
                    string candidate = stem + index.ToString(CultureInfo.InvariantCulture) + ".xml";
                    if (!_package.PartExists(candidate)) return candidate;
                }
            }

            public void ClearMapped() {
                if (Core(create: false) is { } core) {
                    foreach (string name in MappedCoreProperties) RemoveChildren(core, CoreNamespace(name), name);
                }
                if (Application(create: false) is { } application) {
                    foreach (string name in MappedApplicationProperties) {
                        RemoveChildren(application, OfficeXmlNamespaces.ExtendedProperties, name);
                    }
                }
                DeleteCustom();
            }

            private void DeleteCustom() {
                string? part = RelatedPart(LegacyPptPackageView.CustomPropertiesRelationshipType);
                if (part == null) return;
                _package.DeletePart(part);
                string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(null);
                OfficeXmlTree tree = _package.GetRelationships(null);
                int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                while (node != OfficeXmlTree.NoNode) {
                    int next = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship");
                    string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart,
                        tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (string.Equals(target, part, StringComparison.OrdinalIgnoreCase)) tree.Remove(node);
                    node = next;
                }
            }

            private static void RemoveChildren(OfficeXmlTree tree, int namespaceId, string localName) {
                int child = tree.FindChild(tree.Root, namespaceId, localName);
                while (child != OfficeXmlTree.NoNode) {
                    int next = tree.FindNextSibling(child, namespaceId, localName);
                    tree.Remove(child);
                    child = next;
                }
            }

            public void SetCore(string localName, string? value) {
                int namespaceId = CoreNamespace(localName);
                if (value == null) {
                    if (Core(create: false) is { } existing) RemoveChildren(existing, namespaceId, localName);
                    return;
                }
                OfficeXmlTree core = Core(create: true)!;
                int element = core.FindChild(core.Root, namespaceId, localName);
                if (element == OfficeXmlTree.NoNode) {
                    element = core.CreateElement(namespaceId, localName);
                    core.Append(core.Root, element);
                }
                core.SetInnerText(element, value);
            }

            public void SetCoreDate(string localName, DateTime? value) {
                if (!value.HasValue) {
                    SetCore(localName, null);
                    return;
                }
                SetCore(localName, value.Value.ToUniversalTime().ToString(
                    "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture));
                if (localName is not "created" and not "modified") return;
                OfficeXmlTree core = Core(create: false)!;
                try {
                    core.DeclareNamespace(core.Root, "dcterms", "http://purl.org/dc/terms/");
                    core.DeclareNamespace(core.Root, "xsi", XsiUri);
                } catch (InvalidOperationException) {
                    // The root binds the prefix elsewhere; the value keeps the declaration the part already has.
                }
                int element = core.FindChild(core.Root, OfficeXmlNamespaces.DublinCoreTerms, localName);
                core.SetAttribute(element, core.Namespaces.GetId(XsiUri), "type", "dcterms:W3CDTF");
            }

            public void EnsureApplication() => Application(create: true);

            public void SetApplication(string localName, string value) {
                OfficeXmlTree application = Application(create: true)!;
                int element = application.GetOrInsertChild(application.Root,
                    OfficeXmlNamespaces.ExtendedProperties, localName, ApplicationOrder);
                application.SetInnerText(element, value);
            }

            public void SetCustom(IReadOnlyList<CustomPropertyValue> values) {
                DeleteCustom();
                string name = UniquePartName("/docProps/custom.xml");
                OfficeXmlTree custom = _package.CreateXmlPart(name, CustomPropertiesContentType,
                    OfficeXmlNamespaces.CustomProperties, "Properties", 4 + values.Count * 2);
                custom.DeclareNamespace(custom.Root, string.Empty, "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
                custom.DeclareNamespace(custom.Root, "vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                foreach (CustomPropertyValue value in values) {
                    int property = custom.CreateElement(OfficeXmlNamespaces.CustomProperties, "property");
                    custom.SetAttribute(property, OfficeXmlNamespaces.None, "fmtid", CustomPropertiesFormatId);
                    custom.SetAttribute(property, OfficeXmlNamespaces.None, "pid",
                        value.PropertyId.ToString(CultureInfo.InvariantCulture));
                    custom.SetAttribute(property, OfficeXmlNamespaces.None, "name", value.Name);
                    int variant = custom.CreateElement(OfficeXmlNamespaces.DocPropsVTypes, value.Variant);
                    custom.SetInnerText(variant, value.Text);
                    custom.Append(property, variant);
                    custom.Append(custom.Root, property);
                }
                _package.AddRelationship(null, LegacyPptPackageView.CustomPropertiesRelationshipType, name.Substring(1));
            }
        }

        // ---- Documents the SDK holds ----------------------------------------------------------------------------

        private static void ApplyThroughSdk(PowerPointDocument presentation,
            IReadOnlyList<OfficeOlePropertySection>? summarySections,
            IReadOnlyList<OfficeOlePropertySection>? documentSummarySections) {
            var properties = new SdkPropertyWriter(presentation.OpenXmlDocument);
            properties.ClearMapped();
            if (summarySections != null) ApplySummary(properties, summarySections);
            if (documentSummarySections != null) ApplyDocumentSummary(properties, documentSummarySections);
        }

        /// <summary>Writes document properties into a document the SDK holds.</summary>
        private sealed class SdkPropertyWriter : IPropertyTarget {
            private readonly PresentationDocument _document;

            internal SdkPropertyWriter(PresentationDocument document) => _document = document;

            public void ClearMapped() {
                foreach (string name in MappedCoreProperties) SetCore(name, null);
                Ap.Properties? application = _document.ExtendedFilePropertiesPart?
                    .Properties;
                if (application != null) {
                    application.Application = null;
                    application.TotalTime = null;
                    application.PresentationFormat = null;
                    application.Slides = null;
                    application.Notes = null;
                    application.HiddenSlides = null;
                    application.Manager = null;
                    application.Company = null;
                }
                if (_document.CustomFilePropertiesPart != null) {
                    _document.DeletePart(_document.CustomFilePropertiesPart);
                }
            }

            public void SetCore(string localName, string? value) {
                var properties = _document.PackageProperties;
                switch (localName) {
                    case "title": properties.Title = value; break;
                    case "subject": properties.Subject = value; break;
                    case "creator": properties.Creator = value; break;
                    case "keywords": properties.Keywords = value; break;
                    case "description": properties.Description = value; break;
                    case "lastModifiedBy": properties.LastModifiedBy = value; break;
                    case "revision": properties.Revision = value; break;
                    case "category": properties.Category = value; break;
                    case "contentStatus": properties.ContentStatus = value; break;
                    case "contentType": properties.ContentType = value; break;
                    case "identifier": properties.Identifier = value; break;
                    case "language": properties.Language = value; break;
                    case "version": properties.Version = value; break;
                    default: SetCoreDate(localName, ParseCoreDate(value)); break;
                }
            }

            public void SetCoreDate(string localName, DateTime? value) {
                var properties = _document.PackageProperties;
                switch (localName) {
                    case "lastPrinted": properties.LastPrinted = value; break;
                    case "created": properties.Created = value; break;
                    case "modified": properties.Modified = value; break;
                }
            }

            private Ap.Properties Application() {
                ExtendedFilePropertiesPart part = _document
                    .ExtendedFilePropertiesPart
                    ?? _document.AddExtendedFilePropertiesPart();
                return part.Properties ??= new Ap.Properties();
            }

            public void EnsureApplication() => Application();

            public void SetApplication(string localName, string value) {
                Ap.Properties application = Application();
                switch (localName) {
                    case "Application": application.Application = new Ap.Application(value); break;
                    case "TotalTime": application.TotalTime = new Ap.TotalTime(value); break;
                    case "PresentationFormat": application.PresentationFormat = new Ap.PresentationFormat(value); break;
                    case "Slides": application.Slides = new Ap.Slides(value); break;
                    case "Notes": application.Notes = new Ap.Notes(value); break;
                    case "HiddenSlides": application.HiddenSlides = new Ap.HiddenSlides(value); break;
                    case "Manager": application.Manager = new Ap.Manager(value); break;
                    case "Company": application.Company = new Ap.Company(value); break;
                }
            }

            public void SetCustom(IReadOnlyList<CustomPropertyValue> values) {
                var properties = new List<CustomDocumentProperty>(values.Count);
                foreach (CustomPropertyValue value in values) {
                    var property = new CustomDocumentProperty {
                        FormatId = CustomPropertiesFormatId,
                        PropertyId = value.PropertyId,
                        Name = value.Name
                    };
                    switch (value.Variant) {
                        case "lpwstr": property.VTLPWSTR = new VTLPWSTR(value.Text); break;
                        case "bool": property.VTBool = new VTBool(value.Text); break;
                        case "filetime": property.VTFileTime = new VTFileTime(value.Text); break;
                        case "r4": property.VTFloat = new VTFloat(value.Text); break;
                        case "r8": property.VTDouble = new VTDouble(value.Text); break;
                        case "i4": property.VTInt32 = new VTInt32(value.Text); break;
                        case "ui1": property.VTUnsignedByte = new VTUnsignedByte(value.Text); break;
                        case "ui2": property.VTUnsignedShort = new VTUnsignedShort(value.Text); break;
                        case "ui4": property.VTUnsignedInt32 = new VTUnsignedInt32(value.Text); break;
                        case "i8": property.VTInt64 = new VTInt64(value.Text); break;
                        case "ui8": property.VTUnsignedInt64 = new VTUnsignedInt64(value.Text); break;
                        case "blob": property.VTBlob = new VTBlob(value.Text); break;
                    }
                    properties.Add(property);
                }
                CustomFilePropertiesPart part = _document.CustomFilePropertiesPart
                    ?? _document.AddCustomFilePropertiesPart();
                part.Properties = new Properties(properties);
                part.Properties.Save();
            }
        }
    }
}
