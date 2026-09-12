using AngleSharp.Attributes;
using AngleSharp.Js.Attributes;
using AngleSharp.Js.Proxies;
using Jint.Native.Object;
using Jint.Runtime.Descriptors;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace AngleSharp.Js.Cache
{
    static class CreatorCache
    {
        private static readonly ConcurrentDictionary<Type, ConstructorDefinition> _constructorDefinitions = new();
        private static readonly ConcurrentDictionary<Type, EnumLiteralDefinition> _enumLiteralDefinitions = new();
        private static readonly ConcurrentDictionary<Assembly, ISet<String>> _nonEnumTypeNames = new();

        /// <summary>
        /// Gets what is needed to build the constructor object for a type, or null if the
        /// type is not exposed as one. The answer depends on the type alone, so the null
        /// is cached as well - most exported types do not get a constructor.
        /// </summary>
        public static ConstructorDefinition GetConstructorDefinition(this Type type)
        {
            if (!_constructorDefinitions.TryGetValue(type, out var definition))
            {
                var ti = type.GetTypeInfo();
                var names = ti.GetCustomAttributes<DomNameAttribute>()
                    .Select(m => m.OfficialName)
                    .Where(m => m != null)
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();

                if (names.Length > 0 && !ti.IsEnum)
                {
                    var info = ti.DeclaredConstructors.FirstOrDefault(m => m.GetCustomAttributes<DomConstructorAttribute>().Any());
                    definition = new ConstructorDefinition(type, names, info);
                }

                _constructorDefinitions.TryAdd(type, definition);
            }

            return definition;
        }

        /// <summary>
        /// Gets what is needed to publish an enum as a standalone DOM literal object,
        /// or null if the enum should not be exposed that way.
        /// </summary>
        public static EnumLiteralDefinition GetEnumLiteralDefinition(this Type type)
        {
            if (!_enumLiteralDefinitions.TryGetValue(type, out var definition))
            {
                var ti = type.GetTypeInfo();

                if (ti.IsEnum)
                {
                    var name = ti.GetCustomAttributes<DomNameAttribute>(true).FirstOrDefault()?.OfficialName;

                    if (name != null)
                    {
                        var typeNames = GetNonEnumTypeNames(type.GetAssembly());

                        if (!typeNames.Contains(name))
                        {
                            var members = ti.DeclaredFields
                                .Where(m => m.IsLiteral)
                                .Select(m => new EnumLiteralMember(
                                    m.GetCustomAttributes<DomNameAttribute>().FirstOrDefault()?.OfficialName,
                                    m.GetRawConstantValue()))
                                .Where(m => m.Name != null)
                                .ToArray();

                            if (members.Length > 0)
                            {
                                definition = new EnumLiteralDefinition(name, members);
                            }
                        }
                    }
                }

                _enumLiteralDefinitions.TryAdd(type, definition);
            }

            return definition;
        }

        private static ISet<String> GetNonEnumTypeNames(Assembly assembly)
        {
            if (!_nonEnumTypeNames.TryGetValue(assembly, out var names))
            {
                names = new HashSet<String>(StringComparer.OrdinalIgnoreCase);

                foreach (var exportedType in assembly.ExportedTypes)
                {
                    var ti = exportedType.GetTypeInfo();

                    if (ti.IsEnum)
                    {
                        continue;
                    }

                    var name = ti.GetCustomAttributes<DomNameAttribute>(true).FirstOrDefault()?.OfficialName;

                    if (name != null)
                    {
                        names.Add(name);
                    }
                }

                _nonEnumTypeNames.TryAdd(assembly, names);
            }

            return names;
        }

        private static readonly ConcurrentDictionary<Type, Action<EngineInstance, ObjectInstance>> _constructorFunctionActions = new();

        public static Action<EngineInstance, ObjectInstance> GetConstructorFunctionAction(this Type type)
        {
            if (!_constructorFunctionActions.TryGetValue(type, out var action))
            {
                var constructorFunctions = type.GetTypeInfo().GetMethods().Where(m => m.GetCustomAttributes<DomConstructorFunctionAttribute>().Any());

                if (constructorFunctions.Any())
                {
                    action = (engine, obj) =>
                    {
                        foreach (var constructorFunction in constructorFunctions)
                        {
                            var attribute = constructorFunction.GetCustomAttribute<DomConstructorFunctionAttribute>();

                            var constructorFunctionInstance = new DomConstructorFunctionInstance(engine, constructorFunction, attribute.OfficialName);

                            obj.FastSetProperty(attribute.OfficialName, new PropertyDescriptor(constructorFunctionInstance, false, true, false));
                        }
                    };
                }
                else
                {
                    action = (e, o) => { };
                }

                _constructorFunctionActions.TryAdd(type, action);
            }

            return action;
        }

        private static readonly ConcurrentDictionary<Type, Action<EngineInstance, ObjectInstance>> _instanceActions = new();

        public static Action<EngineInstance, ObjectInstance> GetInstanceAction(this Type type)
        {
            if (!_instanceActions.TryGetValue(type, out var action))
            {
                var info = type.GetTypeInfo().DeclaredConstructors.FirstOrDefault(m => m.GetParameters().Length == 0);

                if (info != null)
                {
                    var attributes = type.GetTypeInfo().GetCustomAttributes<DomInstanceAttribute>();
                    action = (engine, obj) =>
                    {
                        foreach (var attribute in attributes)
                        {
                            var instance = info.Invoke(null);

                            if (instance != null)
                            {
                                var node = engine.GetDomNode(instance);
                                obj.FastSetProperty(attribute.Name, new PropertyDescriptor(node, false, true, false));
                            }
                        }
                    };
                }
                else
                {
                    action = (e, o) => { };
                }

                _instanceActions.TryAdd(type, action);
            }

            return action;
        }
    }

    sealed class EnumLiteralDefinition
    {
        public EnumLiteralDefinition(String name, EnumLiteralMember[] members)
        {
            Name = name;
            Members = members;
        }

        public String Name { get; }

        public EnumLiteralMember[] Members { get; }
    }

    readonly struct EnumLiteralMember
    {
        public EnumLiteralMember(String name, Object value)
        {
            Name = name;
            Value = value;
        }

        public String Name { get; }

        public Object Value { get; }
    }

    /// <summary>
    /// Everything the constructor object of a type is built from. The reflection behind it
    /// is the same for every engine, so it is resolved once and kept by
    /// <see cref="CreatorCache"/> - only the object built from it belongs to an engine.
    /// </summary>
    sealed class ConstructorDefinition
    {
        public ConstructorDefinition(Type type, String[] names, ConstructorInfo info)
        {
            Type = type;
            Names = names;
            Info = info;
        }

        /// <summary>
        /// Gets the type the constructor creates instances of.
        /// </summary>
        public Type Type { get; }

        /// <summary>
        /// Gets the names the constructor is exposed under.
        /// </summary>
        public String[] Names { get; }

        /// <summary>
        /// Gets the primary name of the constructor.
        /// </summary>
        public String Name => Names[0];

        /// <summary>
        /// Gets the constructor to invoke, or null if the type cannot be constructed from
        /// script - naming it is still legal, calling it is not.
        /// </summary>
        public ConstructorInfo Info { get; }
    }
}
