using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The parts a binary presentation projects into its scratch presentation. The projection adds masters, layouts,
    /// themes and the other parts through the scratch document's package while the Open XML SDK has not opened it. Some
    /// engine calls the projection makes (slide titles, OLE objects, media, tables, notes) open the SDK; parts needed
    /// after that (sounds, pictures and theme overrides of slides and notes) are written into the saved package bytes,
    /// with the relationship ids the markup already names.
    /// </summary>
    public sealed partial class PowerPointDocument {
        internal const string LegacyThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
        internal const string LegacyThemeOverrideContentType = "application/vnd.openxmlformats-officedocument.themeOverride+xml";
        internal const string LegacySlideLayoutContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
        internal const string LegacySlideMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";
        internal const string LegacyHandoutMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml";
        internal const string LegacyCommentAuthorsContentType = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
        internal const string LegacyCommentsContentType = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
        internal const string LegacyVbaProjectContentType = "application/vnd.ms-office.vbaProject";
        internal const string LegacyPresentationMainContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
        internal const string LegacyMacroPresentationMainContentType = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";

        internal const string LegacyThemeRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme";
        internal const string LegacyThemeOverrideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "themeOverride";
        internal const string LegacySlideLayoutRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideLayout";
        internal const string LegacySlideMasterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideMaster";
        internal const string LegacySlideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide";
        internal const string LegacyNotesMasterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesMaster";
        internal const string LegacyHandoutMasterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "handoutMaster";
        internal const string LegacyCommentAuthorsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "commentAuthors";
        internal const string LegacyCommentsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "comments";
        internal const string LegacyImageRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "image";
        internal const string LegacyAudioRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "audio";
        internal const string LegacyViewPropertiesRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "viewProps";
        internal const string LegacyOfficeDocumentRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "officeDocument";
        internal const string LegacyVbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

        /// <summary>A part the import added: a part of the package, or one written into the saved bytes later.</summary>
        internal readonly struct LegacyPptImportPart {
            internal LegacyPptImportPart(string? partName, int pending) {
                PartName = partName;
                Pending = pending;
            }

            internal string? PartName { get; }
            internal int Pending { get; }
            internal bool IsNone => PartName == null && Pending < 0;
            internal static LegacyPptImportPart None => new(null, -1);
        }

        /// <summary>Adds the parts of a binary import to its scratch presentation; see the type remarks above.</summary>
        internal sealed class LegacyPptImportParts {
            private static readonly ConditionalWeakTable<PowerPointDocument, LegacyPptImportParts> Contexts = new();

            private readonly PowerPointDocument _document;
            private readonly List<PendingPart> _pendingParts = new();
            private readonly List<PendingRelationship> _pendingRelationships = new();
            private readonly Dictionary<string, List<(string ContentType, byte[] Data, string RelationshipId)>> _images =
                new(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, Dictionary<ushort, string>> _pictureBullets =
                new(StringComparer.OrdinalIgnoreCase);
            private Dictionary<string, uint?>? _customShowIds;

            private LegacyPptImportParts(PowerPointDocument document) {
                _document = document;
            }

            /// <summary>The import state of a scratch presentation, created on first use.</summary>
            internal static LegacyPptImportParts For(PowerPointDocument document) =>
                Contexts.GetValue(document ?? throw new ArgumentNullException(nameof(document)), item => new LegacyPptImportParts(item));

            internal PowerPointDocument Document => _document;

            /// <summary>Whether parts are added to the package now, rather than to the saved bytes.</summary>
            internal bool IsPackageHeld => _document.EditablePackage != null;

            /// <summary>The package, while the SDK has not opened the scratch presentation.</summary>
            internal OfficeOpcEditablePackage Package => _document.EditablePackage
                ?? throw new InvalidOperationException("The scratch presentation's package is held by the Open XML SDK now; parts must be added before the projection opens it.");

            internal bool HasPending => _pendingParts.Count > 0 || _pendingRelationships.Count > 0;

            /// <summary>Tables of masters and special masters that wait for the slides, built through the SDK.</summary>
            internal List<LegacyPptLateTable> LateTables { get; } = new();

            /// <summary>Picture bullet relationships by owner part, so each bullet image is stored once per part.</summary>
            internal Dictionary<ushort, string> PictureBulletsOf(PowerPointPartRef owner) {
                if (!_pictureBullets.TryGetValue(owner.Name, out Dictionary<ushort, string>? relationships)) {
                    relationships = new Dictionary<ushort, string>();
                    _pictureBullets.Add(owner.Name, relationships);
                }
                return relationships;
            }

            /// <summary>Custom show ids by name, null for a name that several shows share.</summary>
            internal Dictionary<string, uint?>? CustomShowIds {
                get => _customShowIds;
                set => _customShowIds = value;
            }

            /// <summary>The main presentation part.</summary>
            internal PowerPointPartRef PresentationPart {
                get {
                    if (_document.ReadablePackage is { } package) {
                        string? main = package.GetRelationshipTarget(null, FindRelationshipId(package, null, LegacyOfficeDocumentRelationshipType) ?? string.Empty);
                        if (main != null) return PowerPointPartRef.Package(_document, main);
                        throw new InvalidDataException("The scratch presentation has no main part.");
                    }
                    return PowerPointPartRef.Package(_document, "/ppt/presentation.xml").Current;
                }
            }

            /// <summary>The internal parts a part relates to with a relationship type, in relationship order.</summary>
            internal static List<PowerPointPartRef> Related(PowerPointPartRef owner, string relationshipType) =>
                owner.RelatedParts(relationshipType).ToList();

            /// <summary>The relationship id of the first internal relationship from <paramref name="owner"/> to <paramref name="target"/>.</summary>
            internal string? RelationshipIdTo(PowerPointPartRef owner, PowerPointPartRef target) {
                OfficeOpcEditablePackage package = Package;
                string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner.Name);
                if (!package.PartExists(relationships)) return null;
                OfficeXmlTree tree = package.GetXml(relationships);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                     node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationships, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (resolved != null && string.Equals(resolved, target.Name, StringComparison.OrdinalIgnoreCase)) {
                        return tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                    }
                }
                return null;
            }

            // ---- Creating parts -----------------------------------------------------------------------------------

            /// <summary>
            /// Adds an XML part related from <paramref name="owner"/>, filled by <paramref name="fill"/>, and returns the
            /// relationship id. <paramref name="relationshipId"/> names the id; otherwise the SDK's kind of generated id is used.
            /// </summary>
            internal string AddXmlPart(PowerPointPartRef owner, string relationshipType, string requestedName, string contentType,
                int namespaceId, string rootName, Action<OfficeXmlNode> fill, out LegacyPptImportPart part, string? relationshipId = null,
                bool numbered = true) {
                if (IsPackageHeld) {
                    OfficeOpcEditablePackage package = Package;
                    string name = NameFor(package, contentType, requestedName, numbered);
                    OfficeXmlTree tree = package.CreateXmlPart(name, contentType, namespaceId, rootName, 32);
                    DeclareCommonNamespaces(tree);
                    fill(new OfficeXmlNode(tree, tree.Root));
                    string id = package.AddRelationship(owner.Name, relationshipType, OfficeOpcPartNames.GetRelativeTarget(owner.Name, name),
                        id: relationshipId ?? NewSdkRelationshipId());
                    MarkEdited();
                    part = new LegacyPptImportPart(name, -1);
                    return id;
                }
                OfficeXmlTree standalone = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), namespaceId, rootName, 32);
                DeclareCommonNamespaces(standalone);
                fill(new OfficeXmlNode(standalone, standalone.Root));
                byte[] content = WithDeclaration(standalone.WriteElement(standalone.Root));
                return AddPending(owner, relationshipType, requestedName, contentType, content, numbered, out part);
            }

            /// <summary>Adds an XML part from markup text related from <paramref name="owner"/> and returns the relationship id.</summary>
            internal string AddXmlPart(PowerPointPartRef owner, string relationshipType, string requestedName, string contentType,
                string xml, out LegacyPptImportPart part, string? relationshipId = null, bool numbered = true) {
                byte[] content = WithDeclaration(Encoding.UTF8.GetBytes(xml));
                if (IsPackageHeld) {
                    OfficeOpcEditablePackage package = Package;
                    string name = NameFor(package, contentType, requestedName, numbered);
                    package.CreateBinaryPart(name, contentType, content);
                    string id = package.AddRelationship(owner.Name, relationshipType, OfficeOpcPartNames.GetRelativeTarget(owner.Name, name),
                        id: relationshipId ?? NewSdkRelationshipId());
                    MarkEdited();
                    part = new LegacyPptImportPart(name, -1);
                    return id;
                }
                return AddPending(owner, relationshipType, requestedName, contentType, content, numbered, out part);
            }

            /// <summary>
            /// Adds a binary part related from <paramref name="owner"/> and returns the relationship id. The part is named
            /// <paramref name="requestedName"/>, or numbered on from it when that name is taken or <paramref name="numbered"/>.
            /// </summary>
            internal string AddBinaryPart(PowerPointPartRef owner, string relationshipType, string requestedName, string contentType,
                byte[] data, bool numbered, bool generatedId, out LegacyPptImportPart part) {
                if (IsPackageHeld) {
                    OfficeOpcEditablePackage package = Package;
                    string name = NameFor(package, contentType, requestedName, numbered);
                    CreatePart(package, name, contentType, data);
                    string id = package.AddRelationship(owner.Name, relationshipType, OfficeOpcPartNames.GetRelativeTarget(owner.Name, name),
                        id: generatedId ? NewSdkRelationshipId() : SmallestFreeRelationshipId(package, owner.Name));
                    MarkEdited();
                    part = new LegacyPptImportPart(name, -1);
                    return id;
                }
                return AddPending(owner, relationshipType, requestedName, contentType, data, numbered, out part);
            }

            /// <summary>Relates <paramref name="owner"/> to a part the import added and returns the relationship id.</summary>
            internal string AddRelationship(PowerPointPartRef owner, string relationshipType, LegacyPptImportPart target, bool generatedId) {
                if (IsPackageHeld && target.PartName != null) {
                    OfficeOpcEditablePackage package = Package;
                    string id = package.AddRelationship(owner.Name, relationshipType, OfficeOpcPartNames.GetRelativeTarget(owner.Name, target.PartName),
                        id: generatedId ? NewSdkRelationshipId() : SmallestFreeRelationshipId(package, owner.Name));
                    MarkEdited();
                    return id;
                }
                string pendingId = NewSdkRelationshipId();
                _pendingRelationships.Add(new PendingRelationship(owner.Name, relationshipType, target, pendingId));
                return pendingId;
            }

            /// <summary>
            /// Relates <paramref name="owner"/> to an image, reusing an image part the import already related from it with
            /// the same content when <paramref name="reuseExisting"/>, and returns the relationship id.
            /// </summary>
            internal string AddImage(PowerPointPartRef owner, string contentType, byte[] data, bool reuseExisting) {
                if (!_images.TryGetValue(owner.Name, out List<(string ContentType, byte[] Data, string RelationshipId)>? images)) {
                    images = new List<(string, byte[], string)>();
                    _images.Add(owner.Name, images);
                }
                if (reuseExisting) {
                    foreach ((string existingType, byte[] existingData, string existingId) in images) {
                        if (string.Equals(existingType, contentType, StringComparison.OrdinalIgnoreCase)
                            && existingData.Length == data.Length && existingData.SequenceEqual(data)) {
                            return existingId;
                        }
                    }
                    if (IsPackageHeld && FindPackageImage(owner, contentType, data) is { } packageId) return packageId;
                }
                string extension = ImageExtension(contentType);
                string id = AddBinaryPart(owner, LegacyImageRelationshipType, "/ppt/media/image" + extension, contentType, data,
                    numbered: true, generatedId: true, out _);
                images.Add((contentType, data, id));
                return id;
            }

            // An image part the owner already relates to with the same content, as the SDK finds it among the owner's parts.
            private string? FindPackageImage(PowerPointPartRef owner, string contentType, byte[] data) {
                OfficeOpcEditablePackage package = Package;
                string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner.Name);
                if (!package.PartExists(relationships)) return null;
                OfficeXmlTree tree = package.GetXml(relationships);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                     node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (!string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type"), LegacyImageRelationshipType, StringComparison.Ordinal)
                        || string.Equals(tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationships, tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (target == null || !package.PartExists(target)
                        || !string.Equals(package.GetContentType(target), contentType, StringComparison.OrdinalIgnoreCase)) continue;
                    byte[] existing = package.ReadPartBytes(target);
                    if (existing.Length == data.Length && existing.SequenceEqual(data)) {
                        return tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                    }
                }
                return null;
            }

            private string AddPending(PowerPointPartRef owner, string relationshipType, string requestedName, string contentType,
                byte[] data, bool numbered, out LegacyPptImportPart part) {
                _pendingParts.Add(new PendingPart(requestedName, contentType, data, numbered));
                part = new LegacyPptImportPart(null, _pendingParts.Count - 1);
                string id = NewSdkRelationshipId();
                _pendingRelationships.Add(new PendingRelationship(owner.Name, relationshipType, part, id));
                return id;
            }

            /// <summary>Writes the parts added after the SDK opened the scratch presentation into its saved bytes.</summary>
            internal byte[] ApplyPending(byte[] packageBytes) {
                if (!HasPending) return packageBytes;
                using var input = new MemoryStream(packageBytes, writable: false);
                using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
                var names = new string[_pendingParts.Count];
                for (int index = 0; index < _pendingParts.Count; index++) {
                    PendingPart pending = _pendingParts[index];
                    string name = NameFor(package, pending.ContentType, pending.RequestedName, pending.Numbered);
                    CreatePart(package, name, pending.ContentType, pending.Data);
                    names[index] = name;
                }
                foreach (PendingRelationship relationship in _pendingRelationships) {
                    string target = relationship.Target.PartName ?? names[relationship.Target.Pending];
                    if (!package.PartExists(relationship.Owner)) {
                        throw new InvalidDataException($"The projected part '{relationship.Owner}' is missing from the saved presentation.");
                    }
                    package.AddRelationship(relationship.Owner, relationship.Type,
                        OfficeOpcPartNames.GetRelativeTarget(relationship.Owner, target), id: relationship.Id);
                }
                using var output = new MemoryStream();
                package.Save(output);
                return output.ToArray();
            }

            private void MarkEdited() {
                _document.MarkPackageEdited();
                _document.InvalidatePackageRelationships();
            }

            // ---- Names and ids ------------------------------------------------------------------------------------

            /// <summary>A relationship id the way the SDK generates one for a part added without an id.</summary>
            internal static string NewSdkRelationshipId() => "R" + Guid.NewGuid().ToString("N").Substring(0, 16);

            /// <summary>The smallest free <c>rIdN</c> of a part, as the engine numbers relationships it names itself.</summary>
            internal static string SmallestFreeRelationshipId(OfficeOpcEditablePackage package, string owner) {
                var used = new HashSet<string>(StringComparer.Ordinal);
                string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner);
                if (package.PartExists(relationships)) {
                    OfficeXmlTree tree = package.GetXml(relationships);
                    for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                         node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        if (tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } id) used.Add(id);
                    }
                }
                int next = 1;
                string candidate;
                do {
                    candidate = "rId" + next.ToString(CultureInfo.InvariantCulture);
                    next++;
                } while (used.Contains(candidate));
                return candidate;
            }

            private static string? FindRelationshipId(OfficeOpcEditablePackage package, string? owner, string type) {
                string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner);
                if (!package.PartExists(relationships)) return null;
                OfficeXmlTree tree = package.GetXml(relationships);
                for (int node = tree.FindChild(tree.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                     node = tree.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty), type, StringComparison.Ordinal)) {
                        return tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                    }
                }
                return null;
            }

            private static string NameFor(OfficeOpcEditablePackage package, string contentType, string requestedName, bool numbered) =>
                numbered ? NumberedName(package, contentType, requestedName) : FreeName(package, requestedName);

            // A name numbered on from the parts of the content type, as the SDK numbers new parts: prefix1.ext, prefix2.ext...
            private static string NumberedName(OfficeOpcEditablePackage package, string contentType, string requestedName) {
                SplitName(requestedName, out string stem, out string extension);
                int sequence = package.CountParts(contentType);
                string name;
                do {
                    sequence++;
                    name = stem + sequence.ToString(CultureInfo.InvariantCulture) + extension;
                } while (package.PartExists(name));
                return name;
            }

            private static string FreeName(OfficeOpcEditablePackage package, string requestedName) {
                if (!package.PartExists(requestedName)) return requestedName;
                SplitName(requestedName, out string stem, out string extension);
                int sequence = 1;
                string name;
                do {
                    sequence++;
                    name = stem + sequence.ToString(CultureInfo.InvariantCulture) + extension;
                } while (package.PartExists(name));
                return name;
            }

            private static void SplitName(string name, out string stem, out string extension) {
                int dot = name.LastIndexOf('.');
                bool hasExtension = dot > name.LastIndexOf('/');
                stem = hasExtension ? name.Substring(0, dot) : name;
                extension = hasExtension ? name.Substring(dot) : string.Empty;
                // A requested name may already carry the first number, as /ppt/media/image1.png does.
                int end = stem.Length;
                while (end > 0 && char.IsDigit(stem[end - 1])) end--;
                if (end < stem.Length && end > stem.LastIndexOf('/') + 1) stem = stem.Substring(0, end);
            }

            // Adds a part with its content type: a default for the extension when the package has none, an override when
            // the default names another type.
            private static void CreatePart(OfficeOpcEditablePackage package, string name, string contentType, byte[] data) {
                package.CreateBinaryPart(name, null, data);
                string? declared = package.GetContentType(name);
                int dot = name.LastIndexOf('.');
                if (declared == null && dot > name.LastIndexOf('/') && !contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)) {
                    package.AddDefaultContentType(name.Substring(dot + 1), contentType);
                } else if (!string.Equals(declared, contentType, StringComparison.OrdinalIgnoreCase)) {
                    package.SetOverride(name, contentType);
                }
            }

            private static string ImageExtension(string contentType) => contentType.ToLowerInvariant() switch {
                "image/png" => ".png",
                "image/jpeg" => ".jpeg",
                "image/bmp" => ".bmp",
                "image/tiff" => ".tiff",
                "image/x-emf" => ".emf",
                "image/x-wmf" => ".wmf",
                "image/gif" => ".gif",
                _ => ".bin"
            };

            private static readonly byte[] XmlDeclaration = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>");

            private static byte[] WithDeclaration(byte[] markup) {
                var result = new byte[XmlDeclaration.Length + markup.Length];
                Buffer.BlockCopy(XmlDeclaration, 0, result, 0, XmlDeclaration.Length);
                Buffer.BlockCopy(markup, 0, result, XmlDeclaration.Length, markup.Length);
                return result;
            }

            /// <summary>Declares the a, r and p prefixes on a new part's root, as the SDK writes them.</summary>
            internal static void DeclareCommonNamespaces(OfficeXmlTree tree) {
                tree.DeclareNamespace(tree.Root, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                tree.DeclareNamespace(tree.Root, "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                tree.DeclareNamespace(tree.Root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            }

            private sealed class PendingPart {
                internal PendingPart(string requestedName, string contentType, byte[] data, bool numbered) {
                    RequestedName = requestedName;
                    ContentType = contentType;
                    Data = data;
                    Numbered = numbered;
                }

                internal string RequestedName { get; }
                internal string ContentType { get; }
                internal byte[] Data { get; }
                internal bool Numbered { get; }
            }

            private sealed class PendingRelationship {
                internal PendingRelationship(string owner, string type, LegacyPptImportPart target, string id) {
                    Owner = owner;
                    Type = type;
                    Target = target;
                    Id = id;
                }

                internal string Owner { get; }
                internal string Type { get; }
                internal LegacyPptImportPart Target { get; }
                internal string Id { get; }
            }
        }

        /// <summary>The part names of a part's folder and below, for parts the import names beside their owner.</summary>
        private static string LegacyPartFolder(PowerPointPartRef owner) {
            string name = owner.Name;
            int slash = name.LastIndexOf('/');
            return slash < 0 ? "/" : name.Substring(0, slash + 1);
        }

        /// <summary>Copies the attributes of an element of a package tree, namespace declarations left out, onto another.</summary>
        private static void CopyLegacyAttributes(OfficeXmlNode source, OfficeXmlNode target) {
            if (source.Tree is not { } tree) throw new InvalidOperationException("Attributes are copied from package trees.");
            int count = tree.AttributeCountOf(source.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(source.Node, index, out int namespaceId, out string localName, out string value);
                if (namespaceId == OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeXmlNamespaces.None && localName == "xmlns") continue;
                target.SetAttribute(namespaceId, localName, value);
            }
        }

        /// <summary>The first p:cNvPr below <paramref name="root"/> with this shape id, and the element that owns it.</summary>
        private static OfficeXmlNode FindLegacyShapeById(OfficeXmlNode root, uint shapeId) {
            string id = shapeId.ToString(CultureInfo.InvariantCulture);
            foreach (OfficeXmlNode properties in root.SchemaDescendants(LegacyPptXml.P, "cNvPr")) {
                if (string.Equals(properties.Attribute("id"), id, StringComparison.Ordinal)) {
                    // cNvPr sits in the non-visual properties of its shape.
                    return properties.Parent.Parent;
                }
            }
            return default;
        }
    }
}

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Element building for the binary import: chained attribute setters and schema-ordered children.</summary>
    internal static class LegacyPptImportXml {
        internal static OfficeXmlNode AddP(this OfficeXmlNode parent, string localName) => parent.AppendNew(LegacyPptXml.P, localName);

        internal static OfficeXmlNode AddA(this OfficeXmlNode parent, string localName) => parent.AppendNew(LegacyPptXml.A, localName);

        internal static OfficeXmlNode Set(this OfficeXmlNode node, string localName, string? value) {
            if (value == null) node.RemoveAttribute(localName);
            else node.SetAttribute(localName, value);
            return node;
        }

        internal static OfficeXmlNode Set(this OfficeXmlNode node, string localName, long value) =>
            node.Set(localName, value.ToString(CultureInfo.InvariantCulture));

        internal static OfficeXmlNode Set(this OfficeXmlNode node, string localName, long? value) =>
            node.Set(localName, value?.ToString(CultureInfo.InvariantCulture));

        internal static OfficeXmlNode Set(this OfficeXmlNode node, string localName, bool value) =>
            node.Set(localName, value ? "1" : "0");

        internal static OfficeXmlNode Set(this OfficeXmlNode node, string localName, bool? value) =>
            node.Set(localName, value.HasValue ? value.Value ? "1" : "0" : null);

        internal static OfficeXmlNode SetRelationship(this OfficeXmlNode node, string localName, string? value) {
            if (value == null) node.RemoveAttribute(LegacyPptXml.R, localName);
            else node.SetAttribute(LegacyPptXml.R, localName, value);
            return node;
        }

        /// <summary>
        /// Replaces the child with this name by a new empty one where it stood, or inserts a new one in schema order, as an
        /// SDK typed child setter does, and returns it.
        /// </summary>
        internal static OfficeXmlNode SetChild(this OfficeXmlNode parent, int namespaceId, string localName, OfficeXmlChildOrder order) {
            OfficeXmlNode existing = parent.Child(namespaceId, localName);
            if (existing.IsNone) return parent.InsertOrdered(parent.CreateElement(namespaceId, localName), order);
            OfficeXmlNode added = parent.InsertBefore(parent.CreateElement(namespaceId, localName), existing);
            existing.Remove();
            return added;
        }

        /// <summary>Appends a copy of every child of <paramref name="source"/> under <paramref name="target"/>.</summary>
        internal static void AppendCopiesOfChildren(this OfficeXmlNode target, OfficeXmlNode source) {
            foreach (OfficeXmlNode child in source.Children().ToList()) target.Append(child.Clone());
        }

        /// <summary>Parses markup into a tree of its own and returns its root.</summary>
        internal static OfficeXmlNode Parse(string xml) {
            byte[] bytes = Encoding.UTF8.GetBytes(xml);
            OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
            return new OfficeXmlNode(tree, tree.Root);
        }

        /// <summary>A new element in a tree of its own, for building markup that is inserted later.</summary>
        internal static OfficeXmlNode CreateDetached(int namespaceId, string localName) {
            OfficeXmlTree tree = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), namespaceId, localName, 16);
            return new OfficeXmlNode(tree, tree.Root);
        }

        private static OfficeXmlChildOrder Order(params (int Namespace, string Name, int Rank)[] entries) =>
            new(entries.Select(entry => entry.Namespace).ToArray(), entries.Select(entry => entry.Name).ToArray(),
                entries.Select(entry => entry.Rank).ToArray());

        private static OfficeXmlChildOrder Sequence(int namespaceId, params string[] names) =>
            new(names.Select(_ => namespaceId).ToArray(), names);

        internal static readonly OfficeXmlChildOrder CommonSlideDataOrder = Sequence(LegacyPptXml.P, "bg", "spTree", "custDataLst", "controls", "extLst");
        internal static readonly OfficeXmlChildOrder SlideOrder = Sequence(LegacyPptXml.P, "cSld", "clrMapOvr", "transition", "timing", "extLst");
        internal static readonly OfficeXmlChildOrder SlideLayoutOrder = Sequence(LegacyPptXml.P, "cSld", "clrMapOvr", "transition", "timing", "hf", "extLst");
        internal static readonly OfficeXmlChildOrder SlideMasterOrder = Sequence(LegacyPptXml.P, "cSld", "clrMap", "sldLayoutIdLst", "transition", "timing", "hf", "txStyles", "extLst");
        internal static readonly OfficeXmlChildOrder NotesMasterOrder = Sequence(LegacyPptXml.P, "cSld", "clrMap", "hf", "notesStyle", "extLst");
        internal static readonly OfficeXmlChildOrder HandoutMasterOrder = Sequence(LegacyPptXml.P, "cSld", "clrMap", "hf", "extLst");
        internal static readonly OfficeXmlChildOrder NotesSlideOrder = Sequence(LegacyPptXml.P, "cSld", "clrMapOvr", "extLst");
        internal static readonly OfficeXmlChildOrder ShapeOrder = Sequence(LegacyPptXml.P, "nvSpPr", "spPr", "style", "txBody", "extLst");
        internal static readonly OfficeXmlChildOrder PictureOrder = Sequence(LegacyPptXml.P, "nvPicPr", "blipFill", "spPr", "style", "extLst");
        internal static readonly OfficeXmlChildOrder NonVisualPictureOrder = Sequence(LegacyPptXml.P, "cNvPr", "cNvPicPr", "nvPr");
        internal static readonly OfficeXmlChildOrder NonVisualConnectorOrder = Sequence(LegacyPptXml.P, "cNvPr", "cNvCxnSpPr", "nvPr");
        internal static readonly OfficeXmlChildOrder ApplicationPropertiesOrder = Sequence(LegacyPptXml.P,
            "ph", "audioCd", "wavAudioFile", "audioFile", "videoFile", "quickTimeFile", "custDataLst", "extLst");
        internal static readonly OfficeXmlChildOrder BlipFillOrder = Order(
            (LegacyPptXml.A, "blip", 0), (LegacyPptXml.A, "srcRect", 1), (LegacyPptXml.A, "tile", 2), (LegacyPptXml.A, "stretch", 2));
        internal static readonly OfficeXmlChildOrder ShapePropertiesOrder = Order(
            (LegacyPptXml.A, "xfrm", 0), (LegacyPptXml.A, "custGeom", 1), (LegacyPptXml.A, "prstGeom", 1),
            (LegacyPptXml.A, "noFill", 2), (LegacyPptXml.A, "solidFill", 2), (LegacyPptXml.A, "gradFill", 2), (LegacyPptXml.A, "blipFill", 2),
            (LegacyPptXml.A, "pattFill", 2), (LegacyPptXml.A, "grpFill", 2), (LegacyPptXml.A, "ln", 3), (LegacyPptXml.A, "effectLst", 4),
            (LegacyPptXml.A, "effectDag", 4), (LegacyPptXml.A, "scene3d", 5), (LegacyPptXml.A, "sp3d", 6), (LegacyPptXml.A, "extLst", 7));
        internal static readonly OfficeXmlChildOrder ThemeOrder = Sequence(LegacyPptXml.A, "themeElements", "objectDefaults", "extraClrSchemeLst", "custClrLst", "extLst");
        internal static readonly OfficeXmlChildOrder ThemeElementsOrder = Sequence(LegacyPptXml.A, "clrScheme", "fontScheme", "fmtScheme", "extLst");
    }
}
