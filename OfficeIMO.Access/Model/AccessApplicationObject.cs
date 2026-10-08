namespace OfficeIMO.Access {
    /// <summary>Inert native form, report or action-macro definition. Unknown properties and designer bytes remain preserve-only.</summary>
    public sealed class AccessApplicationObject : AccessNamedObject {
        internal AccessApplicationObject(AccessDocument document, string name, AccessCatalogEntry catalog, string? storagePath) : base(document, name) {
            CatalogEntry = catalog; StoragePath = storagePath;
        }
        /// <summary>Native catalog provenance, including flags and exact catalog row.</summary>
        public AccessCatalogEntry CatalogEntry { get; }
        /// <summary>Qualified application storage path; null means the catalog object has no qualified carrier mapping.</summary>
        public string? StoragePath { get; }
        /// <summary>Qualified designer property tree; null means the definition remains opaque.</summary>
        public AccessDesignerNode? Definition { get; internal set; }
        /// <summary>Qualified action-macro definition; null means the action representation remains preserve-only.</summary>
        public AccessActionMacroInfo? ActionMacro { get; internal set; }
        /// <summary>Object payloads including base, delta, type and unknown streams, without evaluating active content.</summary>
        public IReadOnlyList<AccessStorageStream> Streams { get; internal set; } = Array.AsReadOnly(Array.Empty<AccessStorageStream>());
    }

    /// <summary>A preserved application stream with source storage provenance.</summary>
    public sealed class AccessStorageStream {
        internal AccessStorageStream(string path, byte[] bytes, int? nativeId) { Path = path; NativeId = nativeId; Payload = new AccessOpaqueValue(11, bytes, "Native application payload retained without execution or recompilation."); }
        /// <summary>Native storage path. Some infrastructure names contain a leading control character.</summary>
        public string Path { get; }
        /// <summary>ACE storage-row identity; Jet compound entries have no ACE row identity.</summary>
        public int? NativeId { get; }
        /// <summary>Exact uninterpreted stream bytes, returned defensively.</summary>
        public AccessOpaqueValue Payload { get; }
    }

    /// <summary>A qualified form/report node. Numeric identifiers preserve native property identity without inventing unknown semantics.</summary>
    public sealed class AccessDesignerNode {
        internal AccessDesignerNode(ushort kind, AccessDesignerProperty[] properties, AccessDesignerNode[] children) {
            NativeKind = kind; Properties = Array.AsReadOnly(properties); Children = Array.AsReadOnly(children);
            string? click = properties.FirstOrDefault(x => x.NativeCode == 126)?.Value as string;
            AccessDesignerProperty? embedded = properties.FirstOrDefault(x => x.NativeCode == 491 && x.NativeType == 17);
            EventBindings = click == null ? Array.AsReadOnly(Array.Empty<AccessDesignerEvent>())
                : Array.AsReadOnly(new[] { new AccessDesignerEvent("Click", click, embedded?.Payload,
                    embedded == null ? null : AccessNativeActionMacro.Read(embedded.Payload.GetBytes())) });
        }
        /// <summary>Native object/control kind, including unrecognized values.</summary>
        public ushort NativeKind { get; }
        /// <summary>Persisted properties. Expressions and event bindings remain inert strings or opaque values.</summary>
        public IReadOnlyList<AccessDesignerProperty> Properties { get; }
        /// <summary>Nested sections and controls in native order.</summary>
        public IReadOnlyList<AccessDesignerNode> Children { get; }
        /// <summary>Qualified Click event binding and its available embedded action definition. Other event codes remain native properties.</summary>
        public IReadOnlyList<AccessDesignerEvent> EventBindings { get; }
        /// <summary>Qualified Name property when present in this payload.</summary>
        public string? Name => Properties.FirstOrDefault(x => x.NativeCode == 20)?.Value as string;
        /// <summary>Qualified Caption property when present in this payload.</summary>
        public string? Caption => Properties.FirstOrDefault(x => x.NativeCode == 17)?.Value as string;
        /// <summary>Inert form/report record source. SQL text is not executed or inferred as a resolved dependency.</summary>
        public string? RecordSource => Properties.FirstOrDefault(x => x.NativeCode == 156)?.Value as string;
        /// <summary>Inert control field/expression source.</summary>
        public string? ControlSource => Properties.FirstOrDefault(x => x.NativeCode == 27)?.Value as string;
        /// <summary>Inert combo/list row source.</summary>
        public string? RowSource => Properties.FirstOrDefault(x => x.NativeCode == 91)?.Value as string;
        /// <summary>Inert combo/list row-source kind, such as Table/Query.</summary>
        public string? RowSourceType => Properties.FirstOrDefault(x => x.NativeCode == 93)?.Value as string;
        /// <summary>Persisted control width in twips, when explicitly stored with a qualified type.</summary>
        public int? Width => Integer(150);
        /// <summary>Persisted section/control height in twips, when explicitly stored with a qualified type.</summary>
        public int? Height => Integer(44);
        private int? Integer(ushort code) => Properties.FirstOrDefault(x => x.NativeCode == code)?.Value is short number ? number
            : Properties.FirstOrDefault(x => x.NativeCode == code)?.Value is int value ? value : (int?)null;
    }

    /// <summary>An inert designer event binding. Procedures, expressions and action macros are never called.</summary>
    public sealed class AccessDesignerEvent {
        internal AccessDesignerEvent(string name, string expression, AccessOpaqueValue? payload, AccessActionMacroInfo? macro) {
            Name = name; Expression = expression; NativeEmbeddedMacro = payload; EmbeddedMacro = macro;
        }
        /// <summary>Qualified event name.</summary>
        public string Name { get; }
        /// <summary>Stored expression, procedure or action-macro binding.</summary>
        public string Expression { get; }
        /// <summary>Exact embedded action definition, including unqualified representations.</summary>
        public AccessOpaqueValue? NativeEmbeddedMacro { get; }
        /// <summary>Qualified embedded actions, or null when this event has no qualified embedded action decoder.</summary>
        public AccessActionMacroInfo? EmbeddedMacro { get; }
    }

    /// <summary>An independently qualified inert action-macro definition, distinct from VBA source.</summary>
    public sealed class AccessActionMacroInfo {
        internal AccessActionMacroInfo(string[] actions) { Actions = Array.AsReadOnly(actions); }
        /// <summary>Qualified action names in stored order. Reading or saving never runs them.</summary>
        public IReadOnlyList<string> Actions { get; }
    }

    /// <summary>One native designer property and its exact payload.</summary>
    public sealed class AccessDesignerProperty {
        internal AccessDesignerProperty(uint id, ushort code, uint type, uint width, byte[] bytes, object? value) {
            NativeId = id; NativeCode = code; NativeType = type; NativeDefaultWidth = width;
            Payload = new AccessOpaqueValue(type, bytes, "Exact designer property payload."); Value = value;
        }
        /// <summary>Native record identity, used by base/delta representations.</summary>
        public uint NativeId { get; }
        /// <summary>Native property code.</summary>
        public ushort NativeCode { get; }
        /// <summary>Persisted value type.</summary>
        public uint NativeType { get; }
        /// <summary>Persisted default-width metadata; zero-length defaults are not fabricated as values.</summary>
        public uint NativeDefaultWidth { get; }
        /// <summary>Qualified typed value or an opaque value for unsupported/default representations.</summary>
        public object? Value { get; }
        /// <summary>Exact property payload.</summary>
        public AccessOpaqueValue Payload { get; }
    }
}
