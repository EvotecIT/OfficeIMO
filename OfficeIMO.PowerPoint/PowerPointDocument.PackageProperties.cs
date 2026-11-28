using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string CorePropertiesRelationshipType = OfficeOpcPackage.CorePropertiesRelationshipType;
        private const string ApplicationPropertiesPartName = "/docProps/app.xml";
        private const string SchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        private const string DateTermsNamespace = "http://purl.org/dc/terms/";

        private Dictionary<string, DateTime?>? _packageDateOverrides;

        internal string? ReadBuiltinProperty(int propertyNamespace, string name) =>
            TryGetPackagePropertyRoot(core: true, edit: false, create: false, out OfficeXmlNode root)
                ? ReadDirectPropertyText(root.Child(propertyNamespace, name)) : PowerPointUtils.ReadSdkBuiltinProperty(this, name);

        internal void WriteBuiltinProperty(int propertyNamespace, string name, string? value) {
            if (!TryGetPackagePropertyRoot(core: true, edit: true, create: value != null, out OfficeXmlNode root)) {
                PowerPointUtils.WriteSdkBuiltinProperty(this, name, value);
                return;
            }
            if (root.Exists) {
                SetCoreProperty(root, propertyNamespace, name, value);
                MarkPackageEdited();
            }
        }

        internal DateTime? ReadBuiltinDateProperty(int propertyNamespace, string name) {
            if (TryGetPackagePropertyRoot(core: true, edit: false, create: false, out OfficeXmlNode root)) {
                if (_packageDateOverrides != null && _packageDateOverrides.TryGetValue(name, out DateTime? assigned)) return assigned;
                if (ReadDirectPropertyText(root.Child(propertyNamespace, name)) is not { } text) return null;
                if (TryReadPackageDate(text, out DateTime date)) return date;
            }
            return PowerPointUtils.ReadSdkBuiltinDateProperty(this, name);
        }

        internal void WriteBuiltinDateProperty(int propertyNamespace, string name, DateTime? value) {
            if (!TryGetPackagePropertyRoot(core: true, edit: true, create: value.HasValue, out OfficeXmlNode root)) {
                PowerPointUtils.WriteSdkBuiltinDateProperty(this, name, value);
                return;
            }
            if (!root.Exists) return;
            SetCoreProperty(root, propertyNamespace, name, value.HasValue ? FormatPackageDate(value.Value) : null);
            RememberAssignedPackageDate(name, value);
            MarkPackageEdited();
        }

        // Setters return the original value until the caller reloads the presentation, including its DateTime kind.
        private void RememberAssignedPackageDate(string name, DateTime? value) =>
            (_packageDateOverrides ??= new Dictionary<string, DateTime?>(StringComparer.Ordinal))[name] = value;

        private void RestoreAssignmentsAfterOpeningSdk() {
            Dictionary<string, DateTime?>? assignedDates = _packageDateOverrides;
            _packageDateOverrides = null;
            if (assignedDates == null) return;
            foreach ((string name, DateTime? value) in assignedDates) {
                PowerPointUtils.WriteSdkBuiltinDateProperty(this, name, value);
            }
        }

        private static bool TryReadPackageDate(string text, out DateTime date) {
            date = default;
            if (text.Length < 19 || text[10] != 'T' || text != text.Trim()) return false;
            try {
                date = XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.RoundtripKind);
                if (date.Kind == DateTimeKind.Utc) date = date.ToLocalTime();
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        internal string ReadApplicationProperty(string name) =>
            TryGetPackagePropertyRoot(core: false, edit: false, create: false, out OfficeXmlNode root)
                ? ReadDirectPropertyText(root.Child(OfficeXmlNamespaces.ExtendedProperties, name)) ?? string.Empty
                : PowerPointUtils.ReadSdkApplicationProperty(this, name);

        internal void WriteApplicationProperty(string name, string? value) {
            if (!TryGetPackagePropertyRoot(core: false, edit: true, create: true, out OfficeXmlNode root)) {
                PowerPointUtils.WriteSdkApplicationProperty(this, name, value);
                return;
            }
            SetApplicationProperty(root, name, value);
            MarkPackageEdited();
        }

        private static string? ReadDirectPropertyText(OfficeXmlNode property) => property.Tree?.GetInnerText(property.Node);

        private bool TryGetPackagePropertyRoot(bool core, bool edit, bool create, out OfficeXmlNode root) {
            root = default;
            if (GetPackage(edit) is not { } package) return false;
            if (!TryReadPropertyPartNames(package, out string? applicationPart, out string? corePart)) return false;
            if ((core ? corePart : applicationPart) is { } name) {
                if (GetPackageXml(name) is not { } tree) return false;
                root = new OfficeXmlNode(tree, tree.Root);
                return core ? IsValidCoreProperties(root) : root.Is(OfficeXmlNamespaces.ExtendedProperties, "Properties");
            }

            if (!create) return true;
            if (!core && package.PartExists(ApplicationPropertiesPartName)) return false;
            root = core ? CreateCorePropertiesPart(package) : CreateApplicationPropertiesPart(package);
            return true;
        }

        private static bool TryReadPropertyPartNames(OfficeOpcEditablePackage package, out string? application, out string? core) {
            application = null;
            core = null;
            string part = OfficeOpcPartNames.GetRelationshipsPartName(null);
            if (!package.PartExists(part)) return true;
            OfficeXmlTree tree = package.GetXml(part);
            for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string type = OfficeXmlNamespaces.NormalizeRelationshipType(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty);
                if (type != ExtendedPropertiesRelationshipType && type != CorePropertiesRelationshipType) continue;
                string? target = string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)
                    ? null : OfficeOpcPartNames.ResolveRelationshipTarget(part, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (!RecordPropertyPart(type, target, ref application, ref core)) return false;
            }
            return true;
        }

        private static bool TryReadPropertyPartNames(OfficeOpcPackage package, out string? application, out string? core) {
            application = null;
            core = null;
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(null)) {
                if (relationship.Type != ExtendedPropertiesRelationshipType && relationship.Type != CorePropertiesRelationshipType) continue;
                if (!RecordPropertyPart(relationship.Type, relationship.TargetPartName, ref application, ref core)) return false;
            }
            return true;
        }

        private static bool RecordPropertyPart(string type, string? target, ref string? application, ref string? core) {
            ref string? previous = ref (type == CorePropertiesRelationshipType ? ref core : ref application);
            if (previous != null || target == null) return false;
            previous = target;
            return true;
        }

        private static bool IsValidCoreProperties(OfficeXmlNode root) {
            if (!root.Is(OfficeXmlNamespaces.CoreProperties, "coreProperties") || root.HasAttributes) return false;
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode property in root.Children()) {
                string name = property.LocalName;
                if (!seen.Add(name) || property.FirstChild.Exists) return false;
                switch (property.NamespaceId) {
                    case OfficeXmlNamespaces.DublinCore when name is "creator" or "description" or "identifier" or "language" or "subject" or "title":
                    case OfficeXmlNamespaces.CoreProperties when name is "category" or "contentStatus" or "contentType" or "keywords" or "lastModifiedBy" or "revision" or "version":
                        if (property.HasAttributes) return false;
                        break;
                    case OfficeXmlNamespaces.DublinCoreTerms when name is "created" or "modified":
                        if (!HasCoreDateType(property) || !TryReadPackageDate(ReadDirectPropertyText(property)!, out _)) return false;
                        break;
                    case OfficeXmlNamespaces.CoreProperties when name == "lastPrinted":
                        if (property.HasAttributes || !TryReadPackageDate(ReadDirectPropertyText(property)!, out _)) return false;
                        break;
                    default:
                        return false;
                }
            }
            return true;
        }

        private static bool HasCoreDateType(OfficeXmlNode date) {
            if (date.Attribute(SchemaInstanceNamespace, "type") is not { } type || !type.EndsWith(":W3CDTF", StringComparison.Ordinal)) return false;
            int colon = type.IndexOf(':');
            if (colon <= 0 || type.Length - colon != 7 || date.LookupNamespace(type.Substring(0, colon)) != DateTermsNamespace) return false;
            foreach (var attribute in date.Attributes()) {
                if (attribute.NamespaceUri != SchemaInstanceNamespace || attribute.LocalName != "type") return false;
            }
            return true;
        }

        private static OfficeXmlNode CreateApplicationPropertiesPart(OfficeOpcEditablePackage package) {
            const string contentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
            OfficeXmlTree application = package.CreateXmlPart(ApplicationPropertiesPartName, contentType,
                OfficeXmlNamespaces.ExtendedProperties, "Properties");
            package.AddRelationship(null, ExtendedPropertiesRelationshipType, ApplicationPropertiesPartName);
            return new OfficeXmlNode(application, application.Root);
        }

        private static OfficeXmlNode CreateCorePropertiesPart(OfficeOpcEditablePackage package) {
            const string contentType = "application/vnd.openxmlformats-package.core-properties+xml";
            string name = "/package/services/metadata/core-properties/" + Guid.NewGuid().ToString("N") + ".psmdcp";
            OfficeXmlTree core = package.CreateXmlPart(name, contentType, OfficeXmlNamespaces.CoreProperties, "coreProperties");
            package.AddRelationship(null, CorePropertiesRelationshipType, name);
            return new OfficeXmlNode(core, core.Root);
        }

        private static void EnsureApplicationProperty(OfficeXmlNode root, string name, string value) {
            if (root.Child(OfficeXmlNamespaces.ExtendedProperties, name).IsNone) SetApplicationProperty(root, name, value);
        }

        private static void SetApplicationProperty(OfficeXmlNode root, string name, string? value) {
            OfficeXmlNode child = root.Child(OfficeXmlNamespaces.ExtendedProperties, name);
            if (child.IsNone) {
                child = root.CreateElement(OfficeXmlNamespaces.ExtendedProperties, name);
                // A missing property is inserted after the first recognized property. Unknown markup keeps its place.
                OfficeXmlNode first = root.FirstChild;
                while (first.Exists && !IsApplicationProperty(first)) first = first.NextSibling;
                child = first.Exists ? first.InsertAfterSelf(child) : root.Append(child);
            }
            if (!string.Equals(ReadDirectPropertyText(child), value, StringComparison.Ordinal)) child.SetInnerText(value ?? string.Empty);
        }

        private static bool IsApplicationProperty(OfficeXmlNode property) =>
            property.NamespaceId == OfficeXmlNamespaces.ExtendedProperties && property.LocalName is
                "Template" or "Manager" or "Company" or "Pages" or "Words" or "Characters" or "PresentationFormat" or
                "Lines" or "Paragraphs" or "Slides" or "Notes" or "TotalTime" or "HiddenSlides" or "MMClips" or "ScaleCrop" or
                "HeadingPairs" or "TitlesOfParts" or "LinksUpToDate" or "CharactersWithSpaces" or "SharedDoc" or "HyperlinkBase" or
                "HLinks" or "HyperlinksChanged" or "DigSig" or "Application" or "AppVersion" or "DocSecurity";

        private static void SetCoreProperty(OfficeXmlNode root, int propertyNamespace, string name, string? value) {
            OfficeXmlNode child = root.Child(propertyNamespace, name);
            if (value == null) {
                child.Remove();
                return;
            }

            if (child.IsNone) {
                child = root.Append(root.CreateElement(propertyNamespace, name));
                if (propertyNamespace == OfficeXmlNamespaces.DublinCoreTerms) {
                    if (child.LookupNamespace("dcterms") != DateTermsNamespace) child.DeclareNamespace("dcterms", DateTermsNamespace);
                    if (child.LookupNamespace("xsi") != SchemaInstanceNamespace) child.DeclareNamespace("xsi", SchemaInstanceNamespace);
                    child.SetAttribute(SchemaInstanceNamespace, "type", "dcterms:W3CDTF", "xsi");
                }
            }
            if (!string.Equals(ReadDirectPropertyText(child), value, StringComparison.Ordinal)) child.SetInnerText(value);
        }
    }
}
