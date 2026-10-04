using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    internal CslText MarkFirstAlignedField() => new CslText(Plain,
        "<span data-csl-first-field=\"true\">" + Html + "</span>", Attempted, Rendered);

    private CslText AffixDisplay(string prefix, string suffix) {
        if (!Html.Contains("<div") || prefix.Length == 0 && suffix.Length == 0) return Affix(prefix, suffix);
        CslText Edge(string value, string position) => new CslText(value, value.Length == 0 ? string.Empty :
            "<span data-csl-display-edge=\"" + position + "\">" + Escape(value) + "</span>");
        CslText result = Join(new[] { Edge(prefix, "start"), this, Edge(suffix, "end") }, string.Empty);
        return new CslText(result.Plain, "<span data-csl-display-owner=\"true\">" + result.Html + "</span>", result.Attempted, result.Rendered);
    }

    /// <summary>Projects display columns after layout transformations without changing their formatting ancestry.</summary>
    internal CslText ProjectDisplay(int maximumCharacters, CancellationToken token) {
        if (!Html.Contains("data-csl-first-field") && !Html.Contains("<div")) return this;
        XElement root = CslStyle.ReadXml("<root>" + Html + "</root>", (int)Math.Min(int.MaxValue, (long)maximumCharacters + 13), 1024, token);
        XElement? first = root.Descendants().FirstOrDefault(element => (string?)element.Attribute("data-csl-first-field") == "true");
        XElement[] displays = first == null ? root.Descendants().Where(element => element.Name.LocalName == "div" &&
            !element.Ancestors().Any(parent => parent.Name.LocalName == "div")).ToArray() : Array.Empty<XElement>();
        if (first == null && displays.Length == 0) return this;

        var boundaries = new Dictionary<XElement, int>();
        for (int index = 0; index < displays.Length; index++) boundaries.Add(displays[index], index * 2 + 1);
        int count = first == null ? displays.Length * 2 + 1 : 2;
        Dictionary<XElement, int> edges = DisplayEdgeTargets(root, boundaries, token);
        var fragments = new Dictionary<int, XElement>();
        int current = 0;
        PartitionDisplay(root.Nodes(), fragments, boundaries, edges, first, ref current, token);
        var output = new StringBuilder(Html.Length);
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested();
            bool block = first != null || (index & 1) != 0;
            string display = first == null ? block ? (string?)displays[index / 2].Attribute("class") ?? string.Empty : string.Empty :
                index == 0 ? "csl-left-margin" : "csl-right-inline";
            if (block) AppendDisplayMarkup("<div class=\"" + Escape(display) + "\">", output, maximumCharacters);
            if (fragments.TryGetValue(index, out XElement? fragment))
                WriteDisplayNodes(fragment.Nodes(), output, maximumCharacters, token);
            if (block) AppendDisplayMarkup("</div>", output, maximumCharacters);
        }
        return new CslText(Plain, output.ToString(), Attempted, Rendered);
    }

    // Each ancestor is copied only for the columns that contain its children.
    // Layout emphasis therefore encloses the prefix, body and suffix in each
    // column, while input emphasis still sees the same inherited properties.
    private static void PartitionDisplay(IEnumerable<XNode> nodes, IDictionary<int, XElement> fragments,
        IReadOnlyDictionary<XElement, int> boundaries, IReadOnlyDictionary<XElement, int> edges, XElement? first, ref int current, CancellationToken token) {
        foreach (XNode node in nodes) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) { DisplayFragment(fragments, current).Add(new XText(text.Value)); continue; }
            if (node is not XElement element) continue;
            bool alignedFirst = ReferenceEquals(element, first);
            int boundary = 0;
            if (alignedFirst || boundaries.TryGetValue(element, out boundary)) {
                if (!alignedFirst) current = boundary;
                DisplayFragment(fragments, current);
                PartitionDisplay(element.Nodes(), fragments, boundaries, edges, first, ref current, token);
                current = alignedFirst ? 1 : boundary + 1;
                continue;
            }
            int resume = current;
            bool ownedEdge = edges.TryGetValue(element, out int destination);
            if (ownedEdge) current = destination;
            try {
                if (!element.Nodes().Any()) {
                    DisplayFragment(fragments, current).Add(new XElement(element));
                    continue;
                }
                var children = new Dictionary<int, XElement>();
                PartitionDisplay(element.Nodes(), children, boundaries, edges, first, ref current, token);
                foreach (var child in children) {
                    token.ThrowIfCancellationRequested();
                    DisplayFragment(fragments, child.Key).Add(new XElement(element.Name, element.Attributes(), child.Value.Nodes()));
                }
            } finally {
                if (ownedEdge) current = resume;
            }
        }
    }

    // Preserve ordinary inline fields before, between and after display blocks.
    // Each wrapper owns its first and last display boundary. An affix or quote
    // moves to that boundary only when the wrapper has no ordinary inline text
    // before or after it; unrelated fields outside the wrapper cannot affect it.
    private static Dictionary<XElement, int> DisplayEdgeTargets(XElement root, IReadOnlyDictionary<XElement, int> boundaries, CancellationToken token) {
        var targets = new Dictionary<XElement, int>();
        foreach (XElement owner in root.Descendants().Where(IsDisplayOwner)) {
            token.ThrowIfCancellationRequested();
            XElement[] fields = owner.Descendants().Where(boundaries.ContainsKey).ToArray();
            if (fields.Length == 0) continue;
            int start = boundaries[fields[0]], end = boundaries[fields[fields.Length - 1]];
            var ordinaryText = new HashSet<int>();
            int position = start - 1;
            ReadDisplayFlow(owner.Nodes(), boundaries, ordinaryText, ref position, false, token);
            foreach (XElement edge in owner.Descendants().Where(element => element.Attribute("data-csl-display-edge") != null)) {
                token.ThrowIfCancellationRequested();
                if (!ReferenceEquals(edge.Ancestors().FirstOrDefault(IsDisplayOwner), owner)) continue;
                if ((string?)edge.Attribute("data-csl-display-edge") == "start" && !ordinaryText.Contains(start - 1)) targets[edge] = start;
                else if ((string?)edge.Attribute("data-csl-display-edge") == "end" && !ordinaryText.Contains(end + 1)) targets[edge] = end;
            }
        }
        return targets;
    }

    private static bool IsDisplayOwner(XElement element) => (string?)element.Attribute("data-csl-display-owner") == "true";

    private static void ReadDisplayFlow(IEnumerable<XNode> nodes, IReadOnlyDictionary<XElement, int> boundaries,
        ISet<int> ordinaryText, ref int position, bool edge, CancellationToken token) {
        foreach (XNode node in nodes) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) { if (!edge && text.Value.Length > 0) ordinaryText.Add(position); continue; }
            if (node is not XElement element) continue;
            if (boundaries.TryGetValue(element, out int boundary)) { position = boundary + 1; continue; }
            ReadDisplayFlow(element.Nodes(), boundaries, ordinaryText, ref position,
                edge || element.Attribute("data-csl-display-edge") != null, token);
        }
    }

    private static XElement DisplayFragment(IDictionary<int, XElement> fragments, int index) {
        if (!fragments.TryGetValue(index, out XElement? fragment)) {
            fragment = new XElement("root");
            fragments.Add(index, fragment);
        }
        return fragment;
    }

    private static void WriteDisplayNodes(IEnumerable<XNode> nodes, StringBuilder output, int maximumCharacters, CancellationToken token) {
        foreach (XNode node in nodes) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) { AppendDisplayMarkup(Escape(text.Value), output, maximumCharacters); continue; }
            if (node is not XElement element) continue;
            AppendDisplayMarkup("<" + element.Name.LocalName, output, maximumCharacters);
            foreach (XAttribute attribute in element.Attributes())
                AppendDisplayMarkup(" " + attribute.Name.LocalName + "=\"" + Escape(attribute.Value) + "\"", output, maximumCharacters);
            AppendDisplayMarkup(">", output, maximumCharacters);
            WriteDisplayNodes(element.Nodes(), output, maximumCharacters, token);
            AppendDisplayMarkup("</" + element.Name.LocalName + ">", output, maximumCharacters);
        }
    }

    private static void AppendDisplayMarkup(string value, StringBuilder output, int maximumCharacters) {
        if ((long)output.Length + value.Length > maximumCharacters)
            throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
        output.Append(value);
    }
}
