namespace AngleSharp.Js.Cache
{
    using AngleSharp.Attributes;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    /// <summary>
    /// Maps a CLR type onto the single type whose prototype represents it in JS.
    /// </summary>
    /// <remarks>
    /// An instance is wrapped from an internal concrete class (HtmlDivElement), while the
    /// constructor exposed to scripts is built from the exported interface (IHtmlDivElement).
    /// Left alone the two end up with a prototype each, so neither "instanceof" nor
    /// "Object.getPrototypeOf(div) === HTMLDivElement.prototype" can ever hold. Both sides are
    /// therefore folded onto the class that defines the DOM name - the topmost class carrying
    /// it - which is also the class the prototype chain is built from.
    /// </remarks>
    static class PrototypeTypeCache
    {
        private static readonly ConcurrentDictionary<Assembly, IReadOnlyDictionary<String, Type>> _definingTypes = new();
        private static readonly ConcurrentDictionary<Assembly, IReadOnlyDictionary<String, Type>> _exposedTypes = new();
        private static readonly ConcurrentDictionary<Type, Boolean> _domTypes = new();

        //  AngleSharp hangs its non-DOM infrastructure - the parsers, the browsing context, the
        //  requesters - off the very same EventTarget class the DOM is built on, so that name
        //  alone says nothing about being part of the DOM. Everything the DOM does expose
        //  carries a more specific name of its own.
        private const String EventTargetName = "EventTarget";

        /// <summary>
        /// Gets whether instances of the type are represented by a DOM prototype rather than by
        /// Jint's ordinary CLR wrapper.
        /// </summary>
        /// <remarks>
        /// The name may well be inherited - a "b" element answers true through HTMLElement -
        /// which is exactly what makes the DOM view the right one for it. Unlike
        /// <see cref="GetDomPrototypeType"/> the answer depends on the type alone and not on the
        /// engine's set of libraries, so the cache is process-wide.
        /// </remarks>
        public static Boolean IsDomType(this Type type) =>
            _domTypes.GetOrAdd(type, static current =>
            {
                var name = GetCanonicalName(current);
                return name != null && !String.Equals(name, EventTargetName, StringComparison.Ordinal);
            });

        /// <summary>
        /// Gets what the constructor object of the type a prototype belongs to is built from,
        /// or null if the type is not exposed as one.
        /// </summary>
        /// <remarks>
        /// A prototype is keyed by the class defining the DOM name, but it is the exported
        /// interface next to it that carries the [DomName] - HtmlDivElement has none of its
        /// own, IHtmlDivElement is what names HTMLDivElement - so the class has to be traded
        /// back for the interface first.
        /// </remarks>
        public static ConstructorDefinition GetConstructorDefinition(this Type type, IEnumerable<Assembly> libs)
        {
            var definition = type.GetConstructorDefinition();

            if (definition == null)
            {
                var name = GetCanonicalName(type);

                if (name != null)
                {
                    foreach (var lib in libs)
                    {
                        if (GetExposedTypes(lib).TryGetValue(name, out var exposedType))
                        {
                            return exposedType.GetConstructorDefinition();
                        }
                    }
                }
            }

            return definition;
        }

        public static Type GetDomPrototypeType(this Type type, IEnumerable<Assembly> libs)
        {
            var typeInfo = type.GetTypeInfo();

            //  An enum carries the [DomName] of its owner (NodeType is named "Document"), and the
            //  closed instantiations of a generic are mutually non-assignable, so a member resolved
            //  against one of them cannot be invoked on another (IHtmlCollection<T>). Neither may
            //  be folded onto somebody else's prototype.
            if (typeInfo.IsEnum || typeInfo.IsGenericType)
            {
                return type;
            }

            var name = GetCanonicalName(type);

            if (name != null)
            {
                foreach (var lib in libs)
                {
                    if (GetDefiningTypes(lib).TryGetValue(name, out var definingType))
                    {
                        return definingType;
                    }
                }
            }

            return type;
        }

        /// <summary>
        /// Gets the DOM name a type is represented by, which for the many element classes
        /// carrying no name of their own (HtmlBoldElement, HtmlSemanticElement, ...) is the name
        /// of their nearest named ancestor - just like in a browser, where a "b" element is an
        /// HTMLElement.
        /// </summary>
        private static String GetCanonicalName(Type type)
        {
            var current = type;

            while (current != null)
            {
                var baseType = current.GetTypeInfo().BaseType;
                var name = current.GetOfficialName(baseType);

                if (name != null)
                {
                    return name;
                }

                current = baseType;
            }

            return null;
        }

        private static IReadOnlyDictionary<String, Type> GetDefiningTypes(Assembly assembly) =>
            _definingTypes.GetOrAdd(assembly, CreateDefiningTypes);

        private static IReadOnlyDictionary<String, Type> CreateDefiningTypes(Assembly assembly)
        {
            //  Ordinal on purpose: XmlHttpRequest is named "XMLHttpRequest" while the
            //  RequesterState enum next to it is named "XmlHttpRequest".
            var result = new Dictionary<String, Type>(StringComparer.Ordinal);

            foreach (var type in GetLoadableTypes(assembly))
            {
                var typeInfo = type.GetTypeInfo();

                //  Only a class can define a prototype: the interface is what the DOM exposes,
                //  but the class is what an instance is built from and what its base type - and
                //  hence the prototype chain - is taken from.
                if (!typeInfo.IsClass || typeInfo.IsGenericType)
                {
                    continue;
                }

                var baseType = typeInfo.BaseType;
                var names = type.GetOfficialNames(baseType);

                foreach (var name in names)
                {
                    if (String.Equals(name, GetNameOf(baseType), StringComparison.Ordinal))
                    {
                        //  The base type already defines this very name, so this class is not
                        //  the topmost one carrying it.
                        continue;
                    }

                    //  A name may legitimately be defined twice (col and colgroup are both an
                    //  HTMLTableColElement); share a prototype and keep the first class seen.
                    if (!result.ContainsKey(name))
                    {
                        result[name] = type;
                    }
                }
            }

            return result;
        }

        private static String GetNameOf(Type type) =>
            type?.GetOfficialName(type.GetTypeInfo().BaseType);

        private static IReadOnlyDictionary<String, Type> GetExposedTypes(Assembly assembly) =>
            _exposedTypes.GetOrAdd(assembly, CreateExposedTypes);

        /// <summary>
        /// Collects the types a DOM name is exposed by, which is what
        /// <see cref="EngineExtensions.AddConstructors"/> walks - hence exported types only.
        /// </summary>
        private static IReadOnlyDictionary<String, Type> CreateExposedTypes(Assembly assembly)
        {
            var result = new Dictionary<String, Type>(StringComparer.Ordinal);

            foreach (var type in assembly.ExportedTypes)
            {
                var typeInfo = type.GetTypeInfo();

                if (typeInfo.IsEnum)
                {
                    //  An enum carries the [DomName] of the type owning it, not one of its own.
                    continue;
                }

                var names = typeInfo.GetCustomAttributes<DomNameAttribute>()
                    .Select(m => m.OfficialName)
                    .Where(m => m != null)
                    .Distinct(StringComparer.Ordinal);
                var rank = GetExposureRank(type);

                foreach (var name in names)
                {
                    if (!result.TryGetValue(name, out var existing) || rank > GetExposureRank(existing))
                    {
                        result[name] = type;
                    }
                }
            }

            return result;
        }

        private static Int32 GetExposureRank(Type type)
        {
            var typeInfo = type.GetTypeInfo();

            if (typeInfo.IsClass)
            {
                return 2;
            }

            if (typeInfo.IsInterface)
            {
                return 1;
            }

            return 0;
        }

        private static IEnumerable<Type> GetLoadableTypes(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(m => m != null);
            }
        }
    }
}
