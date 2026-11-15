using System.Globalization;
using System.Text;
namespace OfficeIMO.Json;

/// <summary>A JSON Pointer built one segment at a time and formatted only when a finding needs it.</summary>
internal sealed class JsonPointerPath {
    private readonly JsonPointerPath? _parent;
    private readonly string? _name;
    private readonly int _index;

    private JsonPointerPath(JsonPointerPath? parent, string? name, int index) {
        _parent = parent;
        _name = name;
        _index = index;
    }

    public static JsonPointerPath? Child(JsonPointerPath? parent, string name) => new JsonPointerPath(parent, name, -1);

    public static JsonPointerPath? Child(JsonPointerPath? parent, int index) => new JsonPointerPath(parent, null, index);

    public static string Format(JsonPointerPath? path) {
        if (path is null) return string.Empty;
        var segments = new List<JsonPointerPath>();
        for (JsonPointerPath? p = path; p is not null; p = p._parent) segments.Add(p);
        var text = new StringBuilder();
        for (int i = segments.Count - 1; i >= 0; i--) {
            text.Append('/');
            JsonPointerPath segment = segments[i];
            if (segment._name is null) text.Append(segment._index.ToString(CultureInfo.InvariantCulture));
            else AppendEscaped(text, segment._name);
        }
        return text.ToString();
    }

    /// <summary>Escapes one reference token: <c>~</c> as <c>~0</c> and <c>/</c> as <c>~1</c>.</summary>
    public static void AppendEscaped(StringBuilder text, string token) {
        foreach (char c in token) {
            if (c == '~') text.Append("~0");
            else if (c == '/') text.Append("~1");
            else text.Append(c);
        }
    }

    /// <summary>Splits a JSON Pointer into unescaped tokens; returns false for text that is not a pointer.</summary>
    public static bool TryParse(string pointer, out List<string> tokens) {
        tokens = new List<string>();
        if (pointer.Length == 0) return true;
        if (pointer[0] != '/') return false;
        var token = new StringBuilder();
        for (int i = 1; i <= pointer.Length; i++) {
            if (i == pointer.Length || pointer[i] == '/') {
                tokens.Add(token.ToString());
                token.Clear();
                continue;
            }
            char c = pointer[i];
            if (c == '~') {
                if (i + 1 >= pointer.Length) return false;
                char next = pointer[++i];
                if (next == '0') token.Append('~');
                else if (next == '1') token.Append('/');
                else return false;
            } else {
                token.Append(c);
            }
        }
        return true;
    }
}

/// <summary>URI reference resolution (RFC 3986) for <c>$id</c> and <c>$ref</c>, kept as strings so URNs and relative bases work.</summary>
internal static class JsonSchemaUri {
    /// <summary>Resolves a reference against a base; an empty base leaves relative references as written.</summary>
    public static string Resolve(string baseUri, string reference) {
        if (reference.Length == 0) return StripEmptyFragment(baseUri);
        if (HasScheme(reference)) return Normalize(reference);
        if (reference[0] == '#') return ResourcePart(baseUri) + reference;
        string resource = ResourcePart(baseUri);
        if (resource.Length != 0 && IsHierarchical(resource) && Uri.TryCreate(resource, UriKind.Absolute, out Uri? absolute) && Uri.TryCreate(absolute, reference, out Uri? resolved)) {
            return StripEmptyFragment(resolved.AbsoluteUri);
        }
        // An opaque base such as a URN has no path to resolve against.
        if (HasScheme(resource) && !IsHierarchical(resource)) return reference;
        // Without an absolute base, a relative reference resolves against the base's own relative path.
        if (resource.Length != 0) {
            // A query-only reference keeps the base's path (RFC 3986 section 5.2.2).
            if (reference[0] == '?') {
                int query = resource.IndexOf('?');
                return StripEmptyFragment((query < 0 ? resource : resource.Substring(0, query)) + reference);
            }
            int slash = resource.LastIndexOf('/');
            if (slash >= 0 && reference[0] != '/') return StripEmptyFragment(RemoveDotSegments(resource.Substring(0, slash + 1) + reference));
        }
        return StripEmptyFragment(RemoveDotSegments(reference));
    }

    /// <summary>
    /// Removes <c>.</c> and <c>..</c> path segments (RFC 3986 section 5.2.4) from a relative reference, so
    /// <c>./child.json</c> and <c>child.json</c> name the same resource. Leading <c>..</c> segments that have nothing
    /// to remove stay, as a relative path cannot climb above its unknown base.
    /// </summary>
    internal static string RemoveDotSegments(string reference) {
        int end = reference.IndexOfAny(new[] { '?', '#' });
        string path = end < 0 ? reference : reference.Substring(0, end);
        string rest = end < 0 ? string.Empty : reference.Substring(end);
        if (path.IndexOf('.') < 0) return reference;
        string[] segments = path.Split('/');
        var kept = new List<string>(segments.Length);
        bool rooted = path.Length > 0 && path[0] == '/';
        for (int i = rooted ? 1 : 0; i < segments.Length; i++) {
            string segment = segments[i];
            bool last = i == segments.Length - 1;
            if (segment == ".") {
                if (last) kept.Add(string.Empty);
            } else if (segment == "..") {
                if (kept.Count > 0 && kept[kept.Count - 1] != "..") kept.RemoveAt(kept.Count - 1);
                else if (!rooted) kept.Add("..");
                if (last) kept.Add(string.Empty);
            } else {
                kept.Add(segment);
            }
        }
        return (rooted ? "/" : string.Empty) + string.Join("/", kept) + rest;
    }

    public static string ResourcePart(string uri) {
        int hash = uri.IndexOf('#');
        return hash < 0 ? uri : uri.Substring(0, hash);
    }

    /// <summary>Gets the decoded fragment, without '#'; empty when there is none.</summary>
    public static string Fragment(string uri) {
        int hash = uri.IndexOf('#');
        if (hash < 0 || hash == uri.Length - 1) return string.Empty;
        string fragment = uri.Substring(hash + 1);
        try {
            return Uri.UnescapeDataString(fragment);
        } catch (UriFormatException) {
            return fragment;
        }
    }

    public static bool HasFragment(string uri) {
        int hash = uri.IndexOf('#');
        return hash >= 0 && hash < uri.Length - 1;
    }

    public static bool HasScheme(string text) {
        if (text.Length == 0 || !IsAsciiLetter(text[0])) return false;
        for (int i = 1; i < text.Length; i++) {
            char c = text[i];
            if (c == ':') return true;
            if (!(IsAsciiLetter(c) || (c >= '0' && c <= '9') || c == '+' || c == '-' || c == '.')) return false;
        }
        return false;
    }

    private static bool IsAsciiLetter(char c) => (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');

    private static bool IsHierarchical(string uri) {
        int colon = uri.IndexOf(':');
        return colon > 0 && uri.Length > colon + 1 && uri[colon + 1] == '/';
    }

    private static string StripEmptyFragment(string uri) => uri.Length > 0 && uri[uri.Length - 1] == '#' ? uri.Substring(0, uri.Length - 1) : uri;

    private static string Normalize(string uri) {
        string resource = ResourcePart(uri);
        string fragment = uri.Length > resource.Length ? uri.Substring(resource.Length) : string.Empty;
        if (IsHierarchical(resource) && Uri.TryCreate(resource, UriKind.Absolute, out Uri? parsed)) resource = parsed.AbsoluteUri;
        return StripEmptyFragment(resource + fragment);
    }
}

/// <summary>Annotations collected for one instance location, which <c>unevaluatedItems</c> and <c>unevaluatedProperties</c> read.</summary>
internal sealed class JsonSchemaAnnotations {
    public HashSet<string>? Properties;
    public bool AllProperties;
    public int PrefixItems;
    public bool AllItems;
    public HashSet<int>? ContainedItems;

    public void AddProperty(string name) => (Properties ??= new HashSet<string>(StringComparer.Ordinal)).Add(name);

    public bool IsPropertyEvaluated(string name) => AllProperties || (Properties is not null && Properties.Contains(name));

    public bool IsItemEvaluated(int index) => AllItems || index < PrefixItems || (ContainedItems is not null && ContainedItems.Contains(index));

    public void Merge(JsonSchemaAnnotations other) {
        if (other.AllProperties) AllProperties = true;
        if (other.Properties is not null && !AllProperties) {
            foreach (string name in other.Properties) AddProperty(name);
        }
        if (other.AllItems) AllItems = true;
        if (other.PrefixItems > PrefixItems) PrefixItems = other.PrefixItems;
        if (other.ContainedItems is not null) {
            ContainedItems ??= new HashSet<int>();
            ContainedItems.UnionWith(other.ContainedItems);
        }
    }
}
