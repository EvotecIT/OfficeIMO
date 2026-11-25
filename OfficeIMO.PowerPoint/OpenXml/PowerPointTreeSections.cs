using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint.OpenXml;

/// <summary>
/// Brings the sections of a presentation's editable XML tree in line with its slide list after slides were removed or
/// moved, as the slide model does it: members follow slide order, a slide listed twice stays in its first section, a
/// section whose members all left is dropped while one that was empty stays, unlisted slides join the last section with
/// members, and sections are ordered by their first slide. When the model would create a section, which gets a new
/// random id, the plan declines.
/// </summary>
internal static class PowerPointTreeSections {
    internal const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
    internal const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    /// <summary>A planned sync, applied with <see cref="Apply"/>.</summary>
    internal sealed class Plan {
        internal Plan(OfficeXmlTree tree, int list, int p14) {
            Tree = tree;
            List = list;
            P14 = p14;
        }

        internal OfficeXmlTree Tree { get; }
        internal int List { get; }
        internal int P14 { get; }
        // Each kept section with its member entries (existing nodes) and the ids of slides it gains.
        internal List<(int Section, int Members, List<int> Entries, List<uint> Added)> Sections { get; } = new();
    }

    /// <summary>
    /// Plans the sync for <paramref name="slideIds"/>, the slide list as it will be. Returns false when the model has to
    /// do it; <paramref name="plan"/> is null when the presentation has no section list.
    /// </summary>
    internal static bool TryPlan(OfficeXmlTree presentation, int p14, IReadOnlyList<uint> slideIds, out Plan? plan) {
        plan = null;
        int list = FindSectionList(presentation, p14);
        if (list == OfficeXmlTree.NoNode) return true;
        // With no slides the model removes the section list, or keeps sections that wait for slides.
        if (slideIds.Count == 0) return false;

        var index = new Dictionary<uint, int>();
        for (int position = 0; position < slideIds.Count; position++) index[slideIds[position]] = position;

        var planned = new Plan(presentation, list, p14);
        var assigned = new HashSet<uint>();
        for (int section = presentation.FindChild(list, p14, "section"); section != OfficeXmlTree.NoNode;
             section = presentation.FindNextSibling(section, p14, "section")) {
            int members = presentation.FindChild(section, p14, "sldIdLst");
            // The model adds the missing member list; writing it alike is not worth the case.
            if (members == OfficeXmlTree.NoNode) return false;
            var listed = new List<(int Entry, uint Id)>();
            for (int entry = presentation.FindChild(members, p14, "sldId"); entry != OfficeXmlTree.NoNode;
                 entry = presentation.FindNextSibling(entry, p14, "sldId")) {
                if (!TryReadId(presentation, entry, out uint id)) return false;
                listed.Add((entry, id));
            }
            bool wasEmpty = listed.Count == 0;
            var kept = new List<int>();
            foreach ((int entry, uint id) in listed.Where(item => index.ContainsKey(item.Id)).OrderBy(item => index[item.Id])) {
                if (assigned.Add(id)) kept.Add(entry);
            }
            if (wasEmpty || kept.Count > 0) planned.Sections.Add((section, members, kept, new List<uint>()));
        }
        if (planned.Sections.Count == 0) return false;

        List<uint> unassigned = slideIds.Where(id => !assigned.Contains(id)).ToList();
        if (unassigned.Count > 0) {
            int target = planned.Sections.FindLastIndex(section => section.Entries.Count > 0);
            if (target < 0) return false;
            planned.Sections[target].Added.AddRange(unassigned);
        }

        // An empty section stays before the section after it.
        var start = new int[planned.Sections.Count];
        int following = int.MaxValue;
        for (int position = planned.Sections.Count - 1; position >= 0; position--) {
            int first = FirstSlide(presentation, planned.Sections[position].Entries, planned.Sections[position].Added, index);
            if (first != int.MaxValue) following = first;
            start[position] = following;
        }
        var order = Enumerable.Range(0, planned.Sections.Count).OrderBy(position => start[position]).ToList();
        var sorted = order.Select(position => planned.Sections[position]).ToList();
        planned.Sections.Clear();
        planned.Sections.AddRange(sorted);
        plan = planned;
        return true;
    }

    internal static void Apply(Plan plan) {
        OfficeXmlTree tree = plan.Tree;
        foreach ((int section, int members, List<int> entries, List<uint> added) in plan.Sections) {
            while (tree.FirstChildOf(members) != OfficeXmlTree.NoNode) tree.Remove(tree.FirstChildOf(members));
            foreach (int entry in entries) tree.Append(members, entry);
            foreach (uint id in added) {
                int entry = tree.CreateElement(plan.P14, "sldId");
                tree.SetAttribute(entry, OfficeXmlNamespaces.None, "id", id.ToString(CultureInfo.InvariantCulture));
                tree.Append(members, entry);
            }
        }
        while (tree.FirstChildOf(plan.List) != OfficeXmlTree.NoNode) tree.Remove(tree.FirstChildOf(plan.List));
        foreach ((int section, _, _, _) in plan.Sections) tree.Append(plan.List, section);
    }

    // The first p14:sectionLst in the section extension of the presentation's first extension list.
    private static int FindSectionList(OfficeXmlTree presentation, int p14) {
        int extensions = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "extLst");
        if (extensions == OfficeXmlTree.NoNode) return OfficeXmlTree.NoNode;
        for (int extension = presentation.FindChild(extensions, OfficeXmlNamespaces.PresentationML, "ext"); extension != OfficeXmlTree.NoNode;
             extension = presentation.FindNextSibling(extension, OfficeXmlNamespaces.PresentationML, "ext")) {
            if (string.Equals(presentation.GetAttribute(extension, OfficeXmlNamespaces.None, "uri"), SectionListUri, StringComparison.Ordinal)) {
                return presentation.FindChild(extension, p14, "sectionLst");
            }
        }
        return OfficeXmlTree.NoNode;
    }

    private static int FirstSlide(OfficeXmlTree presentation, List<int> entries, List<uint> added, Dictionary<uint, int> index) {
        int first = int.MaxValue;
        foreach (int entry in entries) {
            if (TryReadId(presentation, entry, out uint id) && index.TryGetValue(id, out int position) && position < first) first = position;
        }
        foreach (uint id in added) {
            if (index.TryGetValue(id, out int position) && position < first) first = position;
        }
        return first;
    }

    internal static bool TryReadId(OfficeXmlTree tree, int element, out uint id) {
        id = 0;
        string? value = tree.GetAttribute(element, OfficeXmlNamespaces.None, "id");
        return value != null && uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out id);
    }
}
