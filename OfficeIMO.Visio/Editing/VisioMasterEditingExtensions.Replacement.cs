using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio;

public static partial class VisioMasterEditingExtensions {
    private static VisioMaster ResolveMaster(VisioPage page, string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Master NameU cannot be empty.", nameof(name));
        return page.OwnerDocument?.TryGetMaster(name, out VisioMaster? master) == true && master != null
            ? master : VisioDocument.CreateBuiltinMaster(name);
    }

    private static VisioMaster ResolveStencilMaster(VisioPage page, VisioStencilShape stencil) {
        if (string.IsNullOrWhiteSpace(stencil.SourcePackagePath)) return ResolveMaster(page, stencil.MasterNameU);
        if (page.OwnerDocument?.TryGetMaster(stencil.MasterNameU, out VisioMaster? master) == true && master?.IsPackageBacked == true) {
            VisioStencilMetadata.EnsureSourcePackageMatches(master, stencil.SourcePackagePath!);
            return master;
        }
        // Import into a detached owner. Its captured font scopes and relationships
        // travel with the master; failed editing must not change the live document.
        var imported = VisioDocument.Create(VisioPackageType.Stencil);
        imported.ImportStencilMasters(stencil.SourcePackagePath!, new[] { stencil.MasterNameU });
        return imported.GetMaster(stencil.MasterNameU);
    }

    private static void ReplaceMasters(VisioPage page, IReadOnlyList<VisioShape> selection, VisioMaster master,
        bool resize, VisioStencilShape? stencil) {
        foreach (VisioShape shape in selection) EnsureShapeBelongsToPage(page, shape);
        var selected = new HashSet<VisioShape>(selection);
        var roots = selection.Distinct().Where(shape => !HasSelectedAncestor(shape, selected)).ToArray();
        if (roots.Length == 0) return;
        if (string.IsNullOrWhiteSpace(master.NameU)) throw new ArgumentException("Master NameU cannot be empty.", nameof(master));
        if (master.IsPackageBacked && page.OwnerDocument?.TryGetMaster(master.NameU, out VisioMaster? existing) == true && existing?.IsPackageBacked == true)
            VisioStencilMetadata.EnsureSourcePackageMatches(existing, master.StencilSourcePackagePath ?? string.Empty);
        if (stencil != null && master.IsPackageBacked) {
            if (string.IsNullOrWhiteSpace(stencil.SourcePackagePath))
                throw new InvalidOperationException("Package-backed replacement masters require their trusted stencil source path.");
            VisioStencilMetadata.EnsureSourcePackageMatches(master, stencil.SourcePackagePath!);
        }
        XNamespace ns = VisioDocument.VisioNamespace;
        if ((master.RawMasterContentXml?.Root?.Element(ns + "Shapes")?.Elements(ns + "Shape").Count() ?? 1) > 1 || master.PreservedAdditionalShapeElements.Count > 0)
            throw new NotSupportedException("Master replacement requires one modeled root shape.");

        var definitions = new Dictionary<VisioShape, VisioShape>();
        var masterIds = VisioDocument.GetMasterShapeIdentifiers(master);
        var addedChildren = new Dictionary<VisioShape, VisioShape[]>();
        var rootMaps = new Dictionary<VisioShape, Dictionary<VisioShape, VisioShape>>();
        var references = new Dictionary<VisioShape, IReadOnlyDictionary<string, string>>();
        foreach (VisioShape root in roots) {
            var mapped = new Dictionary<VisioShape, VisioShape>();
            VisioShape? old = root.MasterShape ?? root.Master?.Shape;
            var oldIds = root.Master == null ? null : VisioDocument.GetMasterShapeIdentifiers(root.Master);
            if (root.Children.Count == 0 && (old == null || old.Children.Count == 0) && master.Shape.Children.Count > 0) {
                var clones = VisioDuplicationExtensions.PrepareReplacementChildren(page, root, master);
                foreach (var pair in clones) mapped.Add(pair.Value, pair.Key);
                addedChildren.Add(root, master.Shape.Children.Select(child => clones[child]).ToArray());
            } else MatchTree(root, old, master.Shape, oldIds, mapped);
            rootMaps.Add(root, mapped);
            foreach (var pair in mapped) definitions.Add(pair.Key, pair.Value);
        }
        var added = new HashSet<VisioShape>(rootMaps.Values.SelectMany(map => map.Keys).Where(shape => !page.AllShapes().Contains(shape)));
        var pageIds = VisioDocument.GetPageElementIdentifiers(page, addedChildren.Values.SelectMany(children => children));
        foreach (var item in rootMaps) {
            var ids = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var pair in item.Value) {
                ids.Add(VisioShapeFormulaReferences.NormalizeSheetId(masterIds[pair.Value.Id]), pageIds[pair.Key.Id]);
            }
            references.Add(item.Key, ids);
            foreach (VisioShape clone in item.Value.Keys.Where(added.Contains)) VisioShapeFormulaReferences.Remap(clone, ids);
        }
        if (definitions.Keys.Any(shape => definitions.Values.Contains(shape)))
            throw new NotSupportedException("Replacement definitions must be separate from the edited page instances.");
        double? width = null, height = null;
        if (resize) {
            VisioMeasurementUnit unit = stencil?.DefaultUnit ?? page.DefaultUnit;
            width = stencil == null ? master.Shape.Width : stencil.DefaultWidth.ToInches(unit);
            height = stencil == null ? master.Shape.Height : stencil.DefaultHeight.ToInches(unit);
        }
        Action applyFrames = VisioShapeResizing.PrepareMasterReplacement(page, roots, definitions, width, height, references, addedChildren);
        var textResolver = new VisioNativeTextStyleResolver(page.OwnerDocument, default);
        var addedTextReferences = new Dictionary<VisioShape, IReadOnlyDictionary<VisioShape, IReadOnlyDictionary<string, string>>>();
        foreach (var item in rootMaps) {
            var origins = item.Value.Values.ToDictionary(definition => definition, definition => references[item.Key]);
            foreach (VisioShape clone in item.Value.Keys.Where(added.Contains)) addedTextReferences[clone] = origins;
        }
        Action[] applyText = definitions.Keys.Select(shape => VisioDocument.PrepareRetainedMasterText(shape, textResolver,
            addedTextReferences.TryGetValue(shape, out var origins) ? origins : GetInheritedTextReferences(shape, pageIds))).ToArray();
        var referencedPoints = new HashSet<VisioShape>(page.Connectors.Where(connector => connector.FromConnectionPoint != null && connector.From != null).Select(connector => connector.From!)
            .Concat(page.Connectors.Where(connector => connector.ToConnectionPoint != null && connector.To != null).Select(connector => connector.To!)));

        // No graph or document mutation occurs until all selected trees, artwork,
        // frames and attached native connector geometry have passed preparation.
        page.OwnerDocument?.RegisterMaster(master);
        if (stencil != null && master.IsPackageBacked) VisioStencilMetadata.Apply(master, stencil, catalogName: null);
        applyFrames();
        foreach (Action apply in applyText) apply();
        foreach (var item in addedChildren) foreach (VisioShape child in item.Value) item.Key.Children.Add(child);
        foreach (var pair in definitions) {
            VisioShape shape = pair.Key, definition = pair.Value;
            shape.Master = master;
            shape.MasterShape = definition;
            shape.MasterShapeId = roots.Contains(shape) ? null : masterIds[definition.Id];
            shape.NameU = roots.Contains(shape) ? master.NameU : definition.NameU ?? definition.Name;
            shape.Type = shape.Children.Count > 0 ? "Group" : definition.Type;
            shape.HasInheritedText = false; // The retained old text now belongs to this instance.
            VisioStencilMetadata.Clear(shape);
            if (!referencedPoints.Contains(shape) && !added.Contains(shape)) shape.ConnectionPoints.Clear();
            shape.ForeignResources.Clear();
            for (int i = shape.PreservedShapeChildren.Count - 1; i >= 0; i--)
                if (shape.PreservedShapeChildren[i].RawElement?.Name.LocalName == "ForeignData") shape.PreservedShapeChildren.RemoveAt(i);
            definition.PersistedId = masterIds[definition.Id];
        }
        foreach (VisioShape shape in page.AllShapes()) shape.PersistedId = pageIds[shape.Id];
        foreach (VisioConnector connector in page.Connectors) connector.PersistedId = pageIds[connector.Id];
        if (stencil != null) foreach (VisioShape root in roots) VisioStencilMetadata.Apply(root, stencil, catalogName: null);
    }

    private static bool HasSelectedAncestor(VisioShape shape, ISet<VisioShape> selected) {
        for (VisioShape? parent = shape.Parent; parent != null; parent = parent.Parent)
            if (selected.Contains(parent)) return true;
        return false;
    }

    private static IReadOnlyDictionary<VisioShape, IReadOnlyDictionary<string, string>> GetInheritedTextReferences(VisioShape shape, IReadOnlyDictionary<string, string> pageIds) {
        var result = new Dictionary<VisioShape, IReadOnlyDictionary<string, string>>();
        if (shape.Master is not VisioMaster master) return result;
        VisioShape root = shape;
        while (root.Parent != null && ReferenceEquals(root.Parent.Master, master)) root = root.Parent;
        var masterIds = VisioDocument.GetMasterShapeIdentifiers(master);
        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        void Visit(VisioShape live) {
            if (ReferenceEquals(live.Master, master) && (live.MasterShape ?? master.Shape) is VisioShape definition) {
                ids[VisioShapeFormulaReferences.NormalizeSheetId(masterIds[definition.Id])] = pageIds[live.Id];
                result[definition] = ids;
            }
            foreach (VisioShape child in live.Children) Visit(child);
        }
        Visit(root);
        return result;
    }

    private static void MatchTree(VisioShape live, VisioShape? old, VisioShape replacement,
        IReadOnlyDictionary<string, string>? oldIds, IDictionary<VisioShape, VisioShape> mapped) {
        if (live.Children.Count != replacement.Children.Count || old != null && live.Children.Count != old.Children.Count)
            throw new NotSupportedException("Master replacement must retain the modeled child tree. Create a new instance for a different hierarchy.");
        mapped.Add(live, replacement);
        var slots = new HashSet<int>();
        for (int i = 0; i < live.Children.Count; i++) {
            VisioShape child = live.Children[i];
            int slot = i;
            if (old != null) {
                slot = -1;
                for (int j = 0; j < old.Children.Count; j++) {
                    VisioShape prior = old.Children[j];
                    if (ReferenceEquals(child.MasterShape, prior) || child.MasterShapeId != null && oldIds != null &&
                        VisioShapeFormulaReferences.NormalizeSheetId(child.MasterShapeId) == VisioShapeFormulaReferences.NormalizeSheetId(oldIds[prior.Id])) {
                        slot = j; break;
                    }
                }
                if (slot < 0) throw new NotSupportedException("Every existing child must still refer to its original master slot.");
            }
            if (!slots.Add(slot)) throw new NotSupportedException("Existing children must have distinct master slots.");
            MatchTree(child, old?.Children[slot], replacement.Children[slot], oldIds, mapped);
        }
    }
}
