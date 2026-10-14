using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using System.Xml;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private long _loadedForeignBytes;
    private int _loadedForeignCount;
    private void CaptureForeignResources(PackagePart owner, XDocument xml, IList<VisioForeignResource> resources) {
        XNamespace ns = VisioNamespace, relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var aliases = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement foreign in xml.Descendants(ns + "ForeignData")) {
            XElement? rel = foreign.Element(ns + "Rel");
            string? id = (string?)rel?.Attribute(relNs + "id");
            if (id == null) continue;
            if (aliases.TryGetValue(id, out string? alias)) { rel!.SetAttributeValue(relNs + "id", alias); continue; }
            PackageRelationship relationship = owner.GetRelationship(id);
            if (relationship.TargetMode != TargetMode.Internal) throw new NotSupportedException("External Visio foreign content is not loaded.");
            string type = relationship.RelationshipType;
            if (!type.EndsWith("/image", StringComparison.Ordinal) && !type.EndsWith("/oleObject", StringComparison.Ordinal))
                throw new NotSupportedException("Unsupported Visio foreign content relationship: " + type);
            PackagePart part = owner.Package.GetPart(PackUriHelper.ResolvePartUri(owner.Uri, relationship.TargetUri));
            if (part.GetRelationships().Any()) throw new NotSupportedException("Foreign content with dependent package parts is outside the preservation profile.");
            if (++_loadedForeignCount > 1024) throw new InvalidDataException("Visio foreign content exceeds 1024 resources.");
            if (_loadedForeignBytes >= 128L * 1024 * 1024) throw new InvalidDataException("Visio foreign content exceeds 128 MiB.");
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            byte[] data = OfficeStreamReader.ReadAllBytes(stream, Math.Min(64L * 1024 * 1024, 128L * 1024 * 1024 - _loadedForeignBytes));
            _loadedForeignBytes += data.LongLength;
            // Content-addressed IDs remain collision-free when callers move shapes between
            // independently loaded documents, without changing IDs on every save.
            using var hash = SHA256.Create();
            string remapped = "rIdForeign" + BitConverter.ToString(hash.ComputeHash(data)).Replace("-", "")
                + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(type + "\n" + part.ContentType))).Replace("-", "");
            resources.Add(new VisioForeignResource { RelationshipId = remapped, RelationshipType = type, ContentType = part.ContentType, Bytes = data });
            aliases.Add(id, remapped); rel!.SetAttributeValue(relNs + "id", remapped);
        }
    }

    private static void BindForeignResources(VisioShape shape, IReadOnlyCollection<VisioForeignResource> resources) {
        var ids = ForeignRelationshipIds(shape.PreservedShapeChildren.Where(entry => entry.RawElement != null).Select(entry => entry.RawElement!));
        shape.ForeignResources.AddRange(resources.Where(resource => ids.Contains(resource.RelationshipId)));
        foreach (VisioShape child in shape.Children) BindForeignResources(child, resources);
    }

    private static IEnumerable<VisioForeignResource> ShapeForeignResources(IEnumerable<VisioShape> shapes) {
        foreach (VisioShape shape in shapes) {
            foreach (VisioForeignResource resource in shape.ForeignResources) yield return resource;
            foreach (VisioForeignResource resource in ShapeForeignResources(shape.Children)) yield return resource;
        }
    }

    private static HashSet<string> ForeignRelationshipIds(IEnumerable<XElement> roots) {
        XNamespace ns = VisioNamespace, relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        return new HashSet<string>(roots.SelectMany(root => root.DescendantsAndSelf(ns + "ForeignData"))
            .Elements(ns + "Rel").Attributes(relNs + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);
    }

    private static void WriteForeignResources(PackagePart owner, IEnumerable<VisioForeignResource> resources) {
        var available = resources.GroupBy(resource => resource.RelationshipId, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        if (available.Count == 0) return;
        // Filter against the actual saved shape tree, including opaque preserved XML.
        XDocument content = LoadPackageXml(owner, "Saved Visio foreign-content references");
        var referenced = ForeignRelationshipIds(content.Elements());
        int index = 1;
        foreach (string id in referenced) {
            if (!available.TryGetValue(id, out VisioForeignResource? resource))
                throw new InvalidDataException("Foreign content resource is unavailable: " + id);
            bool image = resource.RelationshipType.EndsWith("/image", StringComparison.Ordinal);
            string directory = image ? "media" : "embeddings";
            Uri target;
            do { target = new Uri("/visio/" + directory + "/foreign" + index++.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin", UriKind.Relative); }
            while (owner.Package.PartExists(target));
            PackagePart part = owner.Package.CreatePart(target, resource.ContentType);
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) stream.Write(resource.Bytes, 0, resource.Bytes.Length);
            owner.CreateRelationship(PackUriHelper.GetRelativeUri(owner.Uri, target), TargetMode.Internal, resource.RelationshipType, resource.RelationshipId);
        }
    }
}
