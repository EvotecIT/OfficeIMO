namespace AngleSharp.Js.Cache
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using AngleSharp.Text;
    using Jint;
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Runtime;
    using Jint.Runtime.Interop;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    /// <summary>
    /// Holds the member layout of every DOM type the process has met, so that reflecting over a
    /// type tree happens once rather than once per engine.
    /// </summary>
    /// <remarks>
    /// The key is the type together with the engine's library set, because the member set is not
    /// a property of the type alone - loading AngleSharp.Css gives an element members it does not
    /// otherwise have. Shapes are cheap to have several of, so a second configuration simply gets
    /// a second one.
    /// </remarks>
    static class DomShapeCache
    {
        private static readonly ConcurrentDictionary<ShapeKey, DomShape> _shapes = new();
        private static HashSet<String> _reservedNames;

        public static DomShape GetOrCreate(Type type, LibrarySet libs, Engine engine)
        {
            var key = new ShapeKey(type, libs);

            if (_shapes.TryGetValue(key, out var shape))
            {
                return shape;
            }

            //  Built outside GetOrAdd: the factory overload taking state is not available on
            //  netstandard2.0, and a closure per call would be paid by every lookup, hit or not.
            return _shapes.GetOrAdd(key, new DomShapeBuilder(type, libs, GetReservedNames(engine)).Build());
        }

        /// <summary>
        /// The names a member of a DOM type used to be skipped for.
        /// </summary>
        /// <remarks>
        /// Registration used to probe the half-built prototype with HasProperty to decide whether
        /// to skip a member, and at that moment the object's chain was still Object.prototype - so
        /// a DOM member named "toString" was silently dropped. A builder has no object to probe, so
        /// the names are read off the engine instead of copied into this file, which keeps the rule
        /// tied to what Jint's Object.prototype actually carries. They are the same in every
        /// engine, so two threads racing here compute the same set and either answer is right.
        /// </remarks>
        private static HashSet<String> GetReservedNames(Engine engine)
        {
            var reserved = _reservedNames;

            if (reserved == null)
            {
                var keys = engine.Intrinsics.Object.PrototypeObject.GetOwnPropertyKeys(Types.String);
                reserved = new HashSet<String>(StringComparer.Ordinal);

                for (var i = 0; i < keys.Count; i++)
                {
                    reserved.Add(keys[i].ToString());
                }

                _reservedNames = reserved;
            }

            return reserved;
        }

        private readonly struct ShapeKey : IEquatable<ShapeKey>
        {
            private readonly Type _type;
            private readonly LibrarySet _libs;

            public ShapeKey(Type type, LibrarySet libs)
            {
                _type = type;
                _libs = libs;
            }

            public Boolean Equals(ShapeKey other) => _type == other._type && _libs.Equals(other._libs);

            public override Boolean Equals(Object obj) => obj is ShapeKey other && Equals(other);

            public override Int32 GetHashCode() => (_type.GetHashCode() * 397) ^ _libs.GetHashCode();
        }
    }

    /// <summary>
    /// Turns the DOM attributes of one type into the member layout its prototype is built from.
    /// </summary>
    /// <remarks>
    /// The walk is the one the per-engine registration used to do, member for member and in the
    /// same order, because the order decides which of two members of the same name survives.
    /// What changed is where the answer goes: into a declaration that any engine can instantiate,
    /// which forces every member implementation to derive its engine from the receiver rather
    /// than close over one.
    /// </remarks>
    sealed class DomShapeBuilder
    {
        private const String ConstructorName = "constructor";

        private readonly Type _type;
        private readonly Type _baseType;
        private readonly String _name;
        private readonly LibrarySet _libs;
        private readonly HashSet<String> _reserved;
        private readonly Boolean _perRealmMethods;
        private readonly List<String> _order;
        private readonly Dictionary<String, Member> _members;

        private Indexer _numericIndexer;
        private Indexer _stringIndexerGetter;
        private Indexer _stringIndexerSetter;

        public DomShapeBuilder(Type type, LibrarySet libs, HashSet<String> reserved)
        {
            _type = type;
            _baseType = type.GetTypeInfo().BaseType ?? typeof(Object);
            _name = type.GetOfficialName(_baseType);
            _libs = libs;
            _reserved = reserved;
            _order = new List<String>();
            _members = new Dictionary<String, Member>(StringComparer.Ordinal);

            //  The window's members are the one set a script reaches without naming a receiver:
            //  they are copied onto the global object, and a bare "alert('hi')" calls with no
            //  "this" at all. A shared implementation has nothing else to derive its engine from,
            //  so the window's methods keep an engine of their own - see Member.Declare.
            _perRealmMethods = typeof(IWindow).GetTypeInfo().IsAssignableFrom(type.GetTypeInfo());
        }

        public DomShape Build()
        {
            SetAllMembers(_type);
            SetExtensionMembers();

            var constructor = _type.GetConstructorDefinition(_libs.Assemblies);
            var builder = new JsObjectShape.Builder();

            if (_name != null)
            {
                builder.ToStringTag(_name);
            }

            for (var i = 0; i < _order.Count; i++)
            {
                var name = _order[i];

                //  "constructor" was written after the members and overrode one of the same name;
                //  the slot below has taken that role over, so a member claiming it gives way.
                if (constructor != null && String.Equals(name, ConstructorName, StringComparison.Ordinal))
                {
                    continue;
                }

                _members[name].Declare(builder, name, _perRealmMethods);
            }

            if (constructor != null)
            {
                //  The attributes an eagerly written constructor had, and still resolved on first
                //  read: a document names a handful of the types an assembly exposes.
                builder.PerRealmSlot(ConstructorName, CreateConstructor, enumerable: false, writable: true, configurable: true);
            }

            var indexers = _numericIndexer != null || _stringIndexerGetter != null || _stringIndexerSetter != null
                ? new DomIndexers(_numericIndexer, _stringIndexerGetter, _stringIndexerSetter)
                : null;

            return new DomShape(_type, _baseType, _name, builder.Build(), indexers, constructor);
        }

        /// <summary>
        /// Produces the "constructor" value of one engine's prototype. It is the very object
        /// script reads off the window under the type's name, which is what keeps naming a type
        /// and asking one of its instances for its constructor the same answer.
        /// </summary>
        private static JsValue CreateConstructor(ObjectInstance prototype)
        {
            var state = DomPrototypeState.Of(prototype);
            return state != null ? state.GetConstructor(state.Shape.Constructor) : JsValue.Undefined;
        }

        private void SetExtensionMembers()
        {
            if (_name == null)
            {
                return;
            }

            foreach (var type in _libs.Assemblies.GetExtensionTypes(_name))
            {
                var typeInfo = type.GetTypeInfo();
                SetExtensionMethods(typeInfo.DeclaredMethods);
            }
        }

        private void SetAllMembers(Type parentType)
        {
            foreach (var type in parentType.GetTypeTree())
            {
                var typeInfo = type.GetTypeInfo();
                SetNormalProperties(typeInfo.DeclaredProperties);
                SetNormalMethods(typeInfo.DeclaredMethods);
                SetNormalEvents(typeInfo.DeclaredEvents);
            }
        }

        private void SetNormalEvents(IEnumerable<EventInfo> eventInfos)
        {
            foreach (var eventInfo in eventInfos)
            {
                foreach (var m in eventInfo.GetCustomAttributes<DomNameAttribute>())
                {
                    SetEvent(m.OfficialName, eventInfo.AddMethod, eventInfo.RemoveMethod);
                }
            }
        }

        private void SetExtensionMethods(IEnumerable<MethodInfo> methods)
        {
            foreach (var entry in methods.GetExtensions())
            {
                var name = entry.Key;
                var value = entry.Value;

                if (IsDeclared(name))
                {
                    // skip
                }
                else if (value.Adder != null && value.Remover != null)
                {
                    SetEvent(name, value.Adder, value.Remover);
                }
                else if (value.Getter != null || value.Setter != null)
                {
                    SetProperty(name, value.Getter, value.Setter, value.Forward);
                }
                else if (value.Other != null)
                {
                    SetMethod(name, value.Other);
                }
            }
        }

        private void SetNormalProperties(IEnumerable<PropertyInfo> properties)
        {
            foreach (var property in properties)
            {
                var indexParameters = property.GetIndexParameters();
                var accessor = property.GetCustomAttribute<DomAccessorAttribute>()?.Type;
                var putsForward = property.GetCustomAttribute<DomPutForwardsAttribute>();
                var names = property
                    .GetCustomAttributes<DomNameAttribute>()
                    .Select(m => m.OfficialName)
                    .ToArray();

                if (accessor == Accessors.Method)
                {
                    // property decorated with Method accessor, so we need to treat it as a method, not a property

                    if (property.GetMethod == null)
                    {
                        throw new InvalidOperationException("Getter not found.");
                    }

                    foreach (var name in names)
                    {
                        SetMethod(name, property.GetMethod);
                    }

                    // methods were set, so continue with the next property
                    continue;
                }

                var isIndexedAccessor = accessor.HasValue && (accessor.Value & (Accessors.Getter | Accessors.Setter)) != 0;

                if (isIndexedAccessor || Array.Exists(names, m => m.Is("item")))
                {
                    SetIndexer(property, indexParameters);
                }

                foreach (var name in names)
                {
                    SetProperty(name, property.GetMethod, property.SetMethod, putsForward);
                }
            }
        }

        private void SetNormalMethods(IEnumerable<MethodInfo> methods)
        {
            foreach (var method in methods)
            {
                foreach (var m in method.GetCustomAttributes<DomNameAttribute>())
                {
                    SetMethod(m.OfficialName, method);
                }
            }
        }

        private void SetEvent(String name, MethodInfo adder, MethodInfo remover) =>
            Define(name, new EventMember(new DomEventDefinition(adder, remover)));

        private void SetProperty(String name, MethodInfo getter, MethodInfo setter, DomPutForwardsAttribute putsForward) =>
            Define(name, new PropertyMember(getter, setter, putsForward));

        private void SetIndexer(PropertyInfo property, ParameterInfo[] indexParameters)
        {
            if (indexParameters.Length != 1)
            {
                return;
            }

            var getter = property.GetMethod;
            var setter = property.SetMethod;

            if (indexParameters[0].ParameterType == typeof(Int32))
            {
                if (getter != null)
                {
                    _numericIndexer = new Indexer(getter);
                }
            }
            else if (indexParameters[0].ParameterType == typeof(String))
            {
                if (getter != null)
                {
                    _stringIndexerGetter = new Indexer(getter);
                }

                if (setter != null)
                {
                    _stringIndexerSetter = new Indexer(setter);
                }
            }
        }

        private void SetMethod(String name, MethodInfo method)
        {
            //TODO Jint
            // If it already has a property with the given name (usually another method),
            // then convert that method to a two-layer method, which decides which one
            // to pick depending on the number (and probably types) of arguments.
            if (!IsDeclared(name))
            {
                Define(name, new MethodMember(method));
            }
        }

        //  What HasProperty answered while the members were being written onto a half-built
        //  prototype: the names declared so far, plus the ones its chain already carried.
        private Boolean IsDeclared(String name) => _members.ContainsKey(name) || _reserved.Contains(name);

        //  A member replacing one of the same name keeps its position, the way rewriting a
        //  dictionary entry did - the order is what a key enumeration reports.
        private void Define(String name, Member member)
        {
            if (!_members.ContainsKey(name))
            {
                _order.Add(name);
            }

            _members[name] = member;
        }

        /// <summary>
        /// One member of a DOM type, as much of it as is the same for every engine.
        /// </summary>
        private abstract class Member
        {
            public abstract void Declare(JsObjectShape.Builder builder, String name, Boolean perRealm);
        }

        private sealed class MethodMember : Member
        {
            private readonly MethodInfo _method;

            public MethodMember(MethodInfo method)
            {
                _method = method;
            }

            public override void Declare(JsObjectShape.Builder builder, String name, Boolean perRealm)
            {
                var method = _method;

                if (perRealm)
                {
                    //  A window method may be called with no receiver at all, so it is the one
                    //  kind of member that cannot find its engine and has to be handed one.
                    builder.PerRealmSlot(
                        name,
                        prototype => Bind(prototype, name, method),
                        enumerable: false,
                        writable: false,
                        configurable: false);
                }
                else
                {
                    builder.Method(
                        name,
                        (thisObject, arguments) => EngineExtensions.CallShared(method, thisObject, arguments),
                        length: 0,
                        enumerable: false,
                        writable: false,
                        configurable: false);
                }
            }

            private static JsValue Bind(ObjectInstance prototype, String name, MethodInfo method)
            {
                var instance = DomPrototypeState.Of(prototype).Instance;
                return new ClrFunction(instance.Jint, name, (thisObject, arguments) =>
                    instance.Call(method, thisObject, arguments));
            }
        }

        private sealed class PropertyMember : Member
        {
            private readonly MethodInfo _getter;
            private readonly MethodInfo _setter;
            private readonly DomPutForwardsAttribute _putsForward;

            public PropertyMember(MethodInfo getter, MethodInfo setter, DomPutForwardsAttribute putsForward)
            {
                _getter = getter;
                _setter = setter;
                _putsForward = putsForward;
            }

            //  Both halves are declared even when only one accessor exists, because that is what
            //  a pair of ClrFunctions over a null MethodInfo used to be: present, and answering
            //  undefined. An accessor is reached through the object it is read from, so unlike a
            //  method it always has a receiver to derive its engine from - the global object
            //  itself when a script names a window member bare.
            public override void Declare(JsObjectShape.Builder builder, String name, Boolean perRealm)
            {
                var getter = _getter;
                var setter = _setter;
                var putsForward = _putsForward;

                //  The forwarding case is decided here rather than per write: it is rare, and the
                //  ordinary setter is on the hot path for every attribute a script assigns.
                var write = putsForward == null
                    ? (Func<JsValue, JsValue[], JsValue>)((thisObject, arguments) => EngineExtensions.CallShared(setter, thisObject, arguments))
                    : ((thisObject, arguments) => Forward(getter, putsForward, thisObject, arguments));

                builder.Accessor(
                    name,
                    (thisObject, arguments) => EngineExtensions.CallShared(getter, thisObject, arguments),
                    write,
                    enumerable: false,
                    configurable: false);
            }

            private static JsValue Forward(MethodInfo getter, DomPutForwardsAttribute putsForward, JsValue thisObject, JsValue[] arguments)
            {
                var instance = thisObject.GetEngineInstance() ?? throw new JavaScriptException("Illegal invocation.");
                var ep = Array.Empty<Object>();
                var that = thisObject as IDomProxy ?? instance.Window;
                var target = getter.Invoke(that.Value, ep);
                var propName = putsForward.PropertyName;
                var prop = getter.ReturnType
                    .GetInheritedProperties()
                    .FirstOrDefault(m => m.GetCustomAttributes<DomNameAttribute>().Any(n => n.OfficialName.Is(propName)));
                var args = instance.BuildArgs(prop.SetMethod, arguments);
                prop.SetMethod.Invoke(target, args);
                return prop.GetMethod.Invoke(target, ep).ToJsValue(instance);
            }
        }

        private sealed class EventMember : Member
        {
            private readonly DomEventDefinition _definition;

            public EventMember(DomEventDefinition definition)
            {
                _definition = definition;
            }

            public override void Declare(JsObjectShape.Builder builder, String name, Boolean perRealm) =>
                builder.Accessor(name, _definition.GetHandler, _definition.SetHandler, enumerable: false, configurable: false);
        }
    }
}
