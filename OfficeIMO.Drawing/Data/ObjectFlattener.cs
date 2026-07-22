using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeIMO.Drawing {
    /// <summary>
    /// Configuration for flattening objects into key/value pairs where keys represent dotted paths.
    /// </summary>
    public class ObjectFlattenerOptions {
        /// <summary>Property names or dotted paths to expand (descend into) rather than treat as simple values.</summary>
        public List<string> ExpandProperties { get; } = new();
        /// <summary>When true, includes the original object under its path in addition to expanded fields.</summary>
        public bool IncludeFullObjects { get; set; }
        /// <summary>Maximum recursion depth when expanding nested objects.</summary>
        public int MaxDepth { get; set; } = int.MaxValue;
        /// <summary>Header casing strategy for generated column names.</summary>
        public HeaderCase HeaderCase { get; set; } = HeaderCase.Raw;
        /// <summary>Optional prefixes to trim from generated headers.</summary>
        public string[] HeaderPrefixTrimPaths { get; set; } = Array.Empty<string>();
        /// <summary>Optional explicit column whitelist (dotted paths).</summary>
        public string[]? Columns { get; set; }
        /// <summary>Paths to exclude from output.</summary>
        public string[] Ignore { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional include filter using property names or dotted paths. When specified, only matching columns are
        /// kept. Unlike <see cref="Columns"/>, this filter also matches a dotted path by its last property segment
        /// and does not define output order.
        /// </summary>
        public string[] IncludeProperties { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional convenience exclude list using property names or dotted paths. A match occurs when the full dotted path
        /// matches or the last segment (property name) matches case-insensitively. This augments <see cref="Ignore"/> which
        /// treats values as dotted-path prefixes.
        /// </summary>
        public string[] ExcludeProperties { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional list of dotted paths to pin to the front of the generated column order when <see cref="Columns"/> is not specified.
        /// Matching paths keep the specified order; any non-matching paths follow in natural order.
        /// </summary>
        public string[] PinnedFirst { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional list of dotted paths to pin to the end of the generated column order. Processed after <see cref="PinnedFirst"/>.
        /// </summary>
        public string[] PinnedLast { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional per-column priority ordering. Lower numbers appear earlier. Keys can be dotted paths or property names.
        /// When not specified, relative discovery order is preserved. Applied after <see cref="PinnedFirst"/> and before <see cref="PinnedLast"/>.
        /// </summary>
        public Dictionary<string, int> PropertyPriority { get; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>How null values are represented.</summary>
        public NullPolicy NullPolicy { get; set; } = NullPolicy.NullLiteral;
        /// <summary>Per‑path default values used when <see cref="NullPolicy.DefaultValue"/> is selected.</summary>
        public Dictionary<string, object?> DefaultValues { get; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>Per‑path formatting delegates applied to values.</summary>
        public Dictionary<string, Func<object?, object?>> Formatters { get; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>How to handle collections.</summary>
        public CollectionMode CollectionMode { get; set; } = CollectionMode.JoinWith;
        /// <summary>Delimiter used when <see cref="CollectionMode.JoinWith"/> is selected.</summary>
        public string CollectionJoinWith { get; set; } = ",";

        /// <summary>
        /// Maps collection paths (e.g., "ScoreBreakdown") to dynamic columns using element properties.
        /// Example: map path "ScoreBreakdown" with KeyProperty="Name", ValueProperty="Value" to produce columns like
        /// "ScoreBreakdown.HasMX" = 2, "ScoreBreakdown.EffectiveSPFSends" = 2.
        /// </summary>
        public System.Collections.Generic.Dictionary<string, CollectionColumnMapping> CollectionMapColumns { get; } = new(System.StringComparer.OrdinalIgnoreCase);

        // Convenience fluent helpers ------------------------------------------------------------
        /// <summary>Adds or appends to <see cref="PinnedFirst"/> while preserving order and uniqueness.</summary>
        public ObjectFlattenerOptions PinFirst(params string[] paths) {
            if (paths == null || paths.Length == 0) return this;
            var list = new List<string>(PinnedFirst ?? Array.Empty<string>());
            foreach (var p in paths)
                if (!string.IsNullOrWhiteSpace(p) && !list.Contains(p, StringComparer.OrdinalIgnoreCase)) list.Add(p);
            PinnedFirst = list.ToArray();
            return this;
        }

        /// <summary>Adds or appends to <see cref="PinnedLast"/> while preserving order and uniqueness.</summary>
        public ObjectFlattenerOptions PinLast(params string[] paths) {
            if (paths == null || paths.Length == 0) return this;
            var list = new List<string>(PinnedLast ?? Array.Empty<string>());
            foreach (var p in paths)
                if (!string.IsNullOrWhiteSpace(p) && !list.Contains(p, StringComparer.OrdinalIgnoreCase)) list.Add(p);
            PinnedLast = list.ToArray();
            return this;
        }

        /// <summary>Sets priority order 1..N for the provided paths (full path or property name).</summary>
        public ObjectFlattenerOptions PriorityOrder(params string[] paths) {
            if (paths == null || paths.Length == 0) return this;
            for (int i = 0; i < paths.Length; i++) {
                var key = paths[i]; if (string.IsNullOrWhiteSpace(key)) continue;
                PropertyPriority[key] = i + 1;
            }
            return this;
        }

        /// <summary>Appends to <see cref="IncludeProperties"/> while preserving uniqueness.</summary>
        public ObjectFlattenerOptions Include(params string[] properties) {
            if (properties == null || properties.Length == 0) return this;
            var list = new List<string>(IncludeProperties ?? Array.Empty<string>());
            foreach (var p in properties)
                if (!string.IsNullOrWhiteSpace(p) && !list.Contains(p, StringComparer.OrdinalIgnoreCase)) list.Add(p);
            IncludeProperties = list.ToArray();
            return this;
        }

        /// <summary>Appends to <see cref="ExcludeProperties"/> while preserving uniqueness.</summary>
        public ObjectFlattenerOptions Exclude(params string[] properties) {
            if (properties == null || properties.Length == 0) return this;
            var list = new List<string>(ExcludeProperties ?? Array.Empty<string>());
            foreach (var p in properties)
                if (!string.IsNullOrWhiteSpace(p) && !list.Contains(p, StringComparer.OrdinalIgnoreCase)) list.Add(p);
            ExcludeProperties = list.ToArray();
            return this;
        }

        /// <summary>
        /// One-call convenience: set pin-first, priority order, and pin-last.
        /// Any argument may be null; order is applied as PinFirst → PriorityOrder → PinLast.
        /// </summary>
        public ObjectFlattenerOptions Order(string[]? pinFirst = null, string[]? priority = null, string[]? pinLast = null) {
            if (pinFirst != null && pinFirst.Length > 0) PinFirst(pinFirst);
            if (priority != null && priority.Length > 0) PriorityOrder(priority);
            if (pinLast != null && pinLast.Length > 0) PinLast(pinLast);
            return this;
        }
    }

    /// <summary>
    /// Flattens objects to a dictionary of dotted-path keys to values suitable for table generation.
    /// </summary>
    public class ObjectFlattener {
        private static readonly ConcurrentDictionary<Type, ObjectFlattenerProperty[]> _propertyCache = new();
        private static readonly ConcurrentDictionary<CollectionMapAccessorKey, CollectionMapAccessors> _collectionMapAccessorCache = new();
        private static readonly ConcurrentDictionary<Type, FieldInfo[]> _valueTupleFieldCache = new();

        /// <summary>
        /// Flattens <paramref name="item"/> into a dictionary according to <paramref name="opts"/>.
        /// </summary>
        public Dictionary<string, object?> Flatten<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(T item, ObjectFlattenerOptions opts) {
            if (opts == null) throw new ArgumentNullException(nameof(opts));
            if (item == null) {
                return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            }

            object source = item!;
            var result = new Dictionary<string, object?>(GetInitialFlattenCapacity(source.GetType(), opts), StringComparer.OrdinalIgnoreCase);
            FlattenInternal(source, result, string.Empty, 0, opts);

            List<string> selectedPaths = ResolvePaths(result.Keys, opts);
            var selected = new Dictionary<string, object?>(selectedPaths.Count, StringComparer.OrdinalIgnoreCase);
            foreach (string path in selectedPaths) {
                selected[path] = result[path];
            }

            return selected;
        }

        /// <summary>
        /// Computes all reachable dotted paths for a given <paramref name="type"/> under <paramref name="opts"/>.
        /// </summary>
        public List<string> GetPaths([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] Type type, ObjectFlattenerOptions opts) {
            var paths = new List<string>(GetInitialFlattenCapacity(type, opts));
            BuildPaths(type, string.Empty, 0, opts, paths);
            return ResolvePaths(paths, opts);
        }

        /// <summary>
        /// Applies the configured include, exclude, pin, and priority rules to discovered table paths.
        /// </summary>
        public List<string> ResolvePaths(IEnumerable<string> paths, ObjectFlattenerOptions opts) {
            if (paths == null) throw new ArgumentNullException(nameof(paths));
            if (opts == null) throw new ArgumentNullException(nameof(opts));
            List<string> input = paths.Where(path => !string.IsNullOrWhiteSpace(path)).ToList();
            List<string> filtered = ApplySelection(input, opts);
            if (opts.Columns != null && opts.Columns.Length > 0) {
                return ApplyExplicitColumnOrdering(filtered, opts.Columns);
            }

            return ApplyOrdering(filtered, opts);
        }

        private static int GetInitialFlattenCapacity(Type type, ObjectFlattenerOptions opts) {
            if (opts.Columns != null && opts.Columns.Length > 0) {
                return opts.Columns.Length;
            }

            if (IsValueTuple(type)) {
                int fieldCount = GetValueTupleFields(type).Length;
                return fieldCount > 0 ? fieldCount : type.IsGenericType ? type.GetGenericArguments().Length : 0;
            }

            return GetObjectFlattenerProperties(type).Length;
        }

        private static void FlattenInternal(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts) {
            if (depth >= opts.MaxDepth) return;

            var type = obj.GetType();

            // Special-case: ValueTuple (struct tuples) expose public fields (Item1..ItemN) not properties.
            if (IsValueTuple(type)) {
                FlattenValueTuple(obj, dict, prefix, depth, opts);
                return;
            }

            var props = GetObjectFlattenerProperties(type);
            if (opts.ExpandProperties.Count == 0) {
                FlattenInternalWithoutExpansion(obj, dict, prefix, opts, props);
                return;
            }

            foreach (var prop in props) {
                var value = prop.GetValue(obj);
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (ShouldIgnorePath(path, opts.Ignore)) continue;

                bool expand = opts.ExpandProperties.Contains(prop.Name) || opts.ExpandProperties.Contains(path);
                bool isCollection = value is IEnumerable && value is not string;

                if (value == null) {
                    dict[path] = ApplyNullPolicy(path, null, opts);
                    continue;
                }

                if (isCollection) {
                    if (opts.CollectionMapColumns.TryGetValue(path, out var map)) {
                        MapCollectionToColumns(path, (IEnumerable)value, map, dict, opts);
                    } else {
                        dict[path] = HandleCollection(path, (IEnumerable)value, opts);
                    }
                    continue;
                }

                if (!expand || IsSimple(prop.PropertyType)) {
                    dict[path] = ApplyFormatting(path, value, opts);
                    continue;
                }

                if (opts.IncludeFullObjects) {
                    dict[path] = value;
                }

                if (depth + 1 < opts.MaxDepth) {
                    FlattenInternal(value, dict, path, depth + 1, opts);
                }
            }
        }

        private static bool IsValueTuple(Type t) {
            if (!t.IsValueType) return false;
            var n = t.Name;
            if (n.StartsWith("ValueTuple`", StringComparison.Ordinal)) return true;
            // non-generic System.ValueTuple
            return string.Equals(t.FullName, "System.ValueTuple", StringComparison.Ordinal);
        }

        private static void FlattenValueTuple(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts) {
            // ITuple exposes tuple items without reflecting over runtime-generated accessors.
            if (obj is System.Runtime.CompilerServices.ITuple tuple) {
                for (int i = 0; i < tuple.Length; i++) {
                    var path = string.IsNullOrEmpty(prefix) ? $"Item{i + 1}" : $"{prefix}.Item{i + 1}";
                    var val = tuple[i];
                    if (val == null) {
                        dict[path] = ApplyNullPolicy(path, null, opts);
                    } else if (IsSimple(val.GetType())) {
                        dict[path] = ApplyFormatting(path, val, opts);
                    } else {
                        FlattenInternal(val, dict, path, depth + 1, opts);
                    }
                }
                return;
            }

            // Fallback: reflect public instance fields Item1..ItemN
            var fields = GetValueTupleFields(obj.GetType());
            int idx = 1;
            foreach (var f in fields) {
                var path = string.IsNullOrEmpty(prefix) ? $"Item{idx}" : $"{prefix}.Item{idx}";
                var val = f.GetValue(obj);
                if (val == null) {
                    dict[path] = ApplyNullPolicy(path, null, opts);
                } else if (IsSimple(val.GetType())) {
                    dict[path] = ApplyFormatting(path, val, opts);
                } else {
                    FlattenInternal(val, dict, path, depth + 1, opts);
                }
                idx++;
            }
        }

        private static void MapCollectionToColumns(string basePath, IEnumerable enumerable, CollectionColumnMapping map, Dictionary<string, object?> dict, ObjectFlattenerOptions opts) {
            Type? lastType = null;
            CollectionMapAccessors? lastAccessors = null;
            foreach (var item in enumerable) {
                if (item == null) continue;
                Type itemType = item.GetType();
                var accessors = itemType == lastType
                    ? lastAccessors!
                    : GetCollectionMapAccessors(itemType, map);
                lastType = itemType;
                lastAccessors = accessors;
                if (!accessors.IsValid) continue;

                var keyObj = accessors.GetKey(item);
                if (keyObj == null) continue;
                var key = keyObj.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(key)) continue;

                var colPath = basePath + "." + key;
                if (ShouldIgnorePath(colPath, opts.Ignore)) continue;
                var value = accessors.GetValue(item);
                dict[colPath] = ApplyFormatting(colPath, value, opts);
            }
        }

        private static void FlattenInternalWithoutExpansion(object obj, Dictionary<string, object?> dict, string prefix, ObjectFlattenerOptions opts, ObjectFlattenerProperty[] props) {
            foreach (var prop in props) {
                var value = prop.GetValue(obj);
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (ShouldIgnorePath(path, opts.Ignore)) continue;

                if (value == null) {
                    dict[path] = ApplyNullPolicy(path, null, opts);
                    continue;
                }

                if (value is IEnumerable enumerable && value is not string) {
                    if (opts.CollectionMapColumns.TryGetValue(path, out var map)) {
                        MapCollectionToColumns(path, enumerable, map, dict, opts);
                    } else {
                        dict[path] = HandleCollection(path, enumerable, opts);
                    }
                    continue;
                }

                dict[path] = ApplyFormatting(path, value, opts);
            }
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2072",
            Justification = "The public GetPaths(Type) boundary requires public properties and fields, and recursive expansion is explicitly selected by the caller. NativeAOT callers that expand nested models must preserve those nested model members as part of their application contract.")]
        private static void BuildPaths([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] Type type, string prefix, int depth, ObjectFlattenerOptions opts, List<string> paths) {
            if (depth >= opts.MaxDepth) return;
            if (IsValueTuple(type)) {
                // Prefer counting actual Item* fields for precision (covers non-generic System.ValueTuple)
                int itemCount = GetValueTupleFields(type).Length;
                // If field count is 0 but the type is generic, fall back to generic arity (covers ITuple-backed cases)
                if (itemCount == 0 && type.IsGenericType)
                    itemCount = type.GetGenericArguments().Length;

                for (int i = 1; i <= itemCount; i++) {
                    var path = string.IsNullOrEmpty(prefix) ? $"Item{i}" : $"{prefix}.Item{i}";
                    if (!ShouldIgnorePath(path, opts.Ignore))
                        paths.Add(path);
                }
                return;
            }

            var props = GetObjectFlattenerProperties(type);
            if (opts.ExpandProperties.Count == 0) {
                BuildPathsWithoutExpansion(prefix, opts, paths, props);
                return;
            }

            foreach (var prop in props) {
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (ShouldIgnorePath(path, opts.Ignore)) continue;
                bool expand = opts.ExpandProperties.Contains(prop.Name) || opts.ExpandProperties.Contains(path);
                bool isCollection = typeof(IEnumerable).IsAssignableFrom(prop.PropertyType) && prop.PropertyType != typeof(string);
                if (isCollection) {
                    paths.Add(path);
                    continue;
                }
                if (!expand || opts.IncludeFullObjects || IsSimple(prop.PropertyType)) {
                    paths.Add(path);
                }
                if (expand && !IsSimple(prop.PropertyType)) {
                    BuildPaths(prop.PropertyType, path, depth + 1, opts, paths);
                }
            }
        }

        private static FieldInfo[] GetValueTupleFields(Type valueTupleType)
            => _valueTupleFieldCache.GetOrAdd(valueTupleType, CreateValueTupleFields);

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2070",
            Justification = "ValueTuple fields are part of the runtime tuple contract; the generic Flatten<T> and GetPaths(Type) entry points preserve public fields for the supplied tuple type.")]
        private static FieldInfo[] CreateValueTupleFields(Type valueTupleType) {
            try {
                return valueTupleType
                    .GetFields(BindingFlags.Public | BindingFlags.Instance)
                    .Where(f => f.Name.StartsWith("Item", StringComparison.Ordinal))
                    .OrderBy(f => f.Name, StringComparer.Ordinal)
                    .ToArray();
            } catch { return Array.Empty<FieldInfo>(); }
        }

        private static object? HandleCollection(string path, IEnumerable enumerable, ObjectFlattenerOptions opts) {
            if (opts.CollectionMode == CollectionMode.JoinWith) {
                var joined = JoinCollectionValues(enumerable, opts.CollectionJoinWith);
                return ApplyFormatting(path, joined, opts);
            }
            // ExpandRows handled in SheetBuilder
            return enumerable;
        }

        private static string JoinCollectionValues(IEnumerable enumerable, string separator) {
            var enumerator = enumerable.GetEnumerator();
            try {
                if (!enumerator.MoveNext()) {
                    return string.Empty;
                }

                string first = enumerator.Current?.ToString() ?? string.Empty;
                if (!enumerator.MoveNext()) {
                    return first;
                }

                var joined = new StringBuilder(first);
                do {
                    joined.Append(separator);
                    joined.Append(enumerator.Current?.ToString() ?? string.Empty);
                } while (enumerator.MoveNext());

                return joined.ToString();
            } finally {
                (enumerator as IDisposable)?.Dispose();
            }
        }

        private static object? ApplyFormatting(string path, object? value, ObjectFlattenerOptions opts) {
            if (value == null) return ApplyNullPolicy(path, null, opts);
            if (opts.Formatters.TryGetValue(path, out var formatter)) {
                return formatter(value);
            }
            return value;
        }

        private static object? ApplyNullPolicy(string path, object? value, ObjectFlattenerOptions opts) {
            return opts.NullPolicy switch {
                NullPolicy.EmptyString => string.Empty,
                NullPolicy.DefaultValue => opts.DefaultValues.TryGetValue(path, out var v) ? v : null,
                _ => null
            };
        }

        private static bool IsSimple(Type type) {
            return type.IsPrimitive || type.IsEnum || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) || type == typeof(Guid);
        }

        private static bool ShouldIgnorePath(string path, string[] ignorePaths) {
            for (int i = 0; i < ignorePaths.Length; i++) {
                if (path.StartsWith(ignorePaths[i], StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        private static ObjectFlattenerProperty[] GetObjectFlattenerProperties(Type type)
            => _propertyCache.GetOrAdd(type, CreateObjectFlattenerProperties);

        private static CollectionMapAccessors GetCollectionMapAccessors(Type type, CollectionColumnMapping map) {
            var key = new CollectionMapAccessorKey(type, map.KeyProperty, map.ValueProperty);
            return _collectionMapAccessorCache.GetOrAdd(key, CreateCollectionMapAccessors);
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2075",
            Justification = "Collection column mapping is an explicit property-name contract. NativeAOT callers preserve the mapped item model's public properties through the containing generic row model.")]
        private static CollectionMapAccessors CreateCollectionMapAccessors(CollectionMapAccessorKey key) {
            var keyProperty = key.ItemType.GetProperty(key.KeyProperty);
            var valueProperty = key.ItemType.GetProperty(key.ValueProperty);
            if (keyProperty == null || valueProperty == null) {
                return CollectionMapAccessors.Missing;
            }

            return new CollectionMapAccessors(
                CreateObjectFlattenerPropertyGetter(keyProperty),
                CreateObjectFlattenerPropertyGetter(valueProperty));
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2070",
            Justification = "The generic Flatten<T> and GetPaths(Type) entry points preserve public properties for the supplied row type. Reflection is limited to reading those public properties and does not generate code at runtime.")]
        private static ObjectFlattenerProperty[] CreateObjectFlattenerProperties(Type type) {
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .OrderBy(p => p.MetadataToken)
                .ToArray();
            var result = new ObjectFlattenerProperty[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
                result[i] = new ObjectFlattenerProperty(properties[i], CreateObjectFlattenerPropertyGetter(properties[i]));
            }

            return result;
        }

        private static ObjectFlattenerPropertyGetter CreateObjectFlattenerPropertyGetter(PropertyInfo property) {
            return row => property.GetValue(row, null);
        }

        private static void BuildPathsWithoutExpansion(string prefix, ObjectFlattenerOptions opts, List<string> paths, ObjectFlattenerProperty[] props) {
            foreach (var prop in props) {
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (!ShouldIgnorePath(path, opts.Ignore)) {
                    paths.Add(path);
                }
            }
        }

        private static string LastSegment(string path) {
            if (string.IsNullOrEmpty(path)) return path;
            int i = path.LastIndexOf('.');
            return i >= 0 ? path.Substring(i + 1) : path;
        }

        private static bool LastSegmentEquals(string path, string value) {
            if (string.Equals(path, value, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            int start = path.LastIndexOf('.') + 1;
            if (start == 0 || path.Length - start != value.Length) {
                return false;
            }

            return string.Compare(path, start, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }

        private sealed class ObjectFlattenerProperty {
            private readonly ObjectFlattenerPropertyGetter _getter;

            internal ObjectFlattenerProperty(PropertyInfo property, ObjectFlattenerPropertyGetter getter) {
                Name = property.Name;
                PropertyType = property.PropertyType;
                _getter = getter;
            }

            internal string Name { get; }

            internal Type PropertyType { get; }

            internal object? GetValue(object source) => _getter(source);
        }

        private delegate object? ObjectFlattenerPropertyGetter(object row);

        private readonly struct CollectionMapAccessorKey : IEquatable<CollectionMapAccessorKey> {
            internal CollectionMapAccessorKey(Type itemType, string keyProperty, string valueProperty) {
                ItemType = itemType;
                KeyProperty = keyProperty;
                ValueProperty = valueProperty;
            }

            internal Type ItemType { get; }

            internal string KeyProperty { get; }

            internal string ValueProperty { get; }

            public bool Equals(CollectionMapAccessorKey other)
                => ItemType == other.ItemType
                   && string.Equals(KeyProperty, other.KeyProperty, StringComparison.Ordinal)
                   && string.Equals(ValueProperty, other.ValueProperty, StringComparison.Ordinal);

            public override bool Equals(object? obj)
                => obj is CollectionMapAccessorKey other && Equals(other);

            public override int GetHashCode() {
                unchecked {
                    int hash = ItemType.GetHashCode();
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(KeyProperty);
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(ValueProperty);
                    return hash;
                }
            }
        }

        private sealed class CollectionMapAccessors {
            internal static readonly CollectionMapAccessors Missing = new CollectionMapAccessors();

            private readonly ObjectFlattenerPropertyGetter? _keyGetter;
            private readonly ObjectFlattenerPropertyGetter? _valueGetter;

            private CollectionMapAccessors() {
            }

            internal CollectionMapAccessors(ObjectFlattenerPropertyGetter keyGetter, ObjectFlattenerPropertyGetter valueGetter) {
                _keyGetter = keyGetter;
                _valueGetter = valueGetter;
                IsValid = true;
            }

            internal bool IsValid { get; }

            internal object? GetKey(object source) => _keyGetter!(source);

            internal object? GetValue(object source) => _valueGetter!(source);
        }

        private readonly struct OrderedPath {
            internal OrderedPath(string path, int originalIndex, int priority) {
                Path = path;
                OriginalIndex = originalIndex;
                Priority = priority;
            }

            internal string Path { get; }

            internal int OriginalIndex { get; }

            internal int Priority { get; }
        }

        private static List<string> ApplyOrdering(List<string> input, ObjectFlattenerOptions opts) {
            if (input == null || input.Count == 0) return new List<string>();

            if ((opts.PinnedFirst == null || opts.PinnedFirst.Length == 0)
                && (opts.PinnedLast == null || opts.PinnedLast.Length == 0)
                && opts.PropertyPriority.Count == 0) {
                return new List<string>(input);
            }

            string[] pinnedFirst = opts.PinnedFirst ?? Array.Empty<string>();
            string[] pinnedLast = opts.PinnedLast ?? Array.Empty<string>();
            var result = new List<string>(input.Count);
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // 1) PinnedFirst in the given order
            foreach (var pin in pinnedFirst) {
                var match = input.FirstOrDefault(p => LastSegmentEquals(p, pin));
                if (!string.IsNullOrEmpty(match) && set.Add(match)) result.Add(match);
            }

            // 2) Remaining, grouped by priority ascending
            int GetPriority(string path) {
                if (opts.PropertyPriority.Count == 0) {
                    return 0;
                }

                string segment = LastSegment(path);
                bool hasPath = opts.PropertyPriority.TryGetValue(path, out int pathPriority);
                int segmentPriority = 0;
                bool hasSegment = !string.Equals(path, segment, StringComparison.Ordinal)
                    && opts.PropertyPriority.TryGetValue(segment, out segmentPriority);

                if (hasPath && hasSegment && pathPriority != segmentPriority) {
                    foreach (var priority in opts.PropertyPriority) {
                        if (string.Equals(priority.Key, path, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(priority.Key, segment, StringComparison.OrdinalIgnoreCase)) {
                            return priority.Value;
                        }
                    }
                }

                if (hasPath) return pathPriority;
                if (hasSegment) return segmentPriority;
                return 0;
            }

            var remaining = new List<OrderedPath>(input.Count);
            for (int i = 0; i < input.Count; i++) {
                string path = input[i];
                if (pinnedFirst.Length != 0 && set.Contains(path)) {
                    continue;
                }

                remaining.Add(new OrderedPath(path, i, GetPriority(path)));
            }

            remaining.Sort((left, right) => {
                int priorityComparison = left.Priority.CompareTo(right.Priority);
                return priorityComparison != 0
                    ? priorityComparison
                    : left.OriginalIndex.CompareTo(right.OriginalIndex);
            });

            // 3) PinnedLast moved to the end in the given order
            if (pinnedLast.Length == 0) {
                foreach (var ordered in remaining) if (set.Add(ordered.Path)) result.Add(ordered.Path);
                return result;
            }

            var pinnedLastMatches = new List<string>(pinnedLast.Length);
            foreach (var pin in pinnedLast) {
                foreach (var ordered in remaining) {
                    string path = ordered.Path;
                    if (LastSegmentEquals(path, pin)) {
                        pinnedLastMatches.Add(path);
                        break;
                    }
                }
            }

            if (pinnedLastMatches.Count == 0) {
                foreach (var ordered in remaining) if (set.Add(ordered.Path)) result.Add(ordered.Path);
                return result;
            }

            // Remove pinned-last matches from prioritized
            var pinnedLastSet = new HashSet<string>(pinnedLastMatches, StringComparer.OrdinalIgnoreCase);

            // Merge
            foreach (var ordered in remaining) if (!pinnedLastSet.Contains(ordered.Path) && set.Add(ordered.Path)) result.Add(ordered.Path);
            foreach (var p in pinnedLastMatches) if (set.Add(p)) result.Add(p);

            return result;
        }

        private static List<string> ApplySelection(List<string> input, ObjectFlattenerOptions opts) {
            if (input == null || input.Count == 0) return new List<string>();

            if (opts.Ignore.Length == 0
                && opts.ExcludeProperties.Length == 0
                && opts.IncludeProperties.Length == 0
                && (opts.Columns == null || opts.Columns.Length == 0)) {
                return new List<string>(input);
            }

            HashSet<string>? columns = opts.Columns != null && opts.Columns.Length > 0
                ? new HashSet<string>(opts.Columns, StringComparer.OrdinalIgnoreCase)
                : null;
            HashSet<string>? exclude = opts.ExcludeProperties.Length > 0
                ? new HashSet<string>(opts.ExcludeProperties, StringComparer.OrdinalIgnoreCase)
                : null;
            HashSet<string>? include = opts.IncludeProperties.Length > 0
                ? new HashSet<string>(opts.IncludeProperties, StringComparer.OrdinalIgnoreCase)
                : null;
            var filtered = new List<string>(input.Count);
            foreach (string path in input) {
                if (ShouldIgnorePath(path, opts.Ignore)) {
                    continue;
                }
                if (columns != null && !columns.Contains(path)) {
                    continue;
                }

                string? segment = null;
                if (exclude != null) {
                    segment = LastSegment(path);
                    if (exclude.Contains(path) || exclude.Contains(segment)) {
                        continue;
                    }
                }

                if (include != null) {
                    segment ??= LastSegment(path);
                    if (!include.Contains(path) && !include.Contains(segment)) {
                        continue;
                    }
                }

                filtered.Add(path);
            }

            return filtered;
        }

        private static List<string> ApplyExplicitColumnOrdering(List<string> input, string[] columns) {
            var result = new List<string>(Math.Min(input.Count, columns.Length));
            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string column in columns) {
                string? match = input.FirstOrDefault(path => string.Equals(path, column, StringComparison.OrdinalIgnoreCase));
                if (match != null && added.Add(match)) {
                    result.Add(match);
                }
            }

            return result;
        }
    }
}
namespace OfficeIMO.Drawing {
/// <summary>
/// Configuration for mapping a collection of objects into dynamic columns.
/// </summary>
public sealed class CollectionColumnMapping {
    /// <summary>Property name on the item to use as the header suffix (key).</summary>
    public string KeyProperty { get; set; } = "Name";
    /// <summary>Property name on the item to use as the cell value.</summary>
    public string ValueProperty { get; set; } = "Value";
    /// <summary>Optional prefix for generated column headers.</summary>
    public string? HeaderPrefix { get; set; }
}
}
