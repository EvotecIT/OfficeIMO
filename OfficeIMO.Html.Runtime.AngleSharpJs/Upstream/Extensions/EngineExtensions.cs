namespace AngleSharp.Js
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using AngleSharp.Js.Cache;
    using Jint;
    using Jint.Native;
    using Jint.Native.Number;
    using Jint.Native.Object;
    using Jint.Runtime;
    using Jint.Runtime.Descriptors;
    using Jint.Runtime.Interop;
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    static class EngineExtensions
    {
        public static JsValue ToJsValue(this Object obj, EngineInstance engine)
        {
            if (obj != null)
            {
                if (obj is String)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(String));
                }
                else if (obj is Int32)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(Int32));
                }
                else if (obj is UInt32)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(UInt32));
                }
                else if (obj is Double)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(Double));
                }
                else if (obj is Single)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(Single));
                }
                else if (obj is Boolean)
                {
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(Boolean));
                }
                else if (obj is Enum)
                {
                    switch (obj)
                    {
                        case DocumentReadyState _:
                            var name = ((Enum)obj).GetOfficialName();
                            if (name != null)
                            {
                                return JsValue.FromObjectWithType(engine.Jint, name, typeof(String));
                            }
                            break;
                    }
                    return JsValue.FromObjectWithType(engine.Jint, obj, typeof(Enum));
                }

                return engine.GetDomNode(obj);
            }

            return JsValue.Null;
        }

        public static ClrFunction AsValue(this Engine engine, string name, Func<JsValue, JsValue[], JsValue> func) =>
            new ClrFunction(engine, name, func);

        public static PropertyDescriptor AsProperty(this Engine engine, JsValue getter = null, JsValue setter = null) =>
            new GetSetPropertyDescriptor(getter, setter, true, getter != null && setter != null);

        public static Object[] BuildArgs(this EngineInstance context, MethodBase method, JsValue[] arguments)
        {
            var description = MethodDescription.Of(method);
            var parameters = description.Parameters;
            var initDict = description.InitDict;
            var max = parameters.Length;
            var args = new Object[max];
            var offset = 0;

            if (description.TakesWindow)
            {
                if (arguments.Length == 0 || arguments[0].FromJsValue() is IWindow == false)
                {
                    args[offset++] = context.Window.Value;
                }
            }

            if (description.TakesParamArray)
            {
                max--;
            }

            if (initDict != null && arguments.Length + offset > initDict.Offset)
            {
                arguments = ExpandInitDict(arguments, parameters, initDict, max, offset);
            }

            var n = Math.Min(arguments.Length, max - offset);

            for (var i = 0; i < n; i++)
            {
                var parameter = parameters[i + offset];

                if (parameter.IsOptional && arguments[i].IsUndefined())
                {
                    args[i + offset] = parameter.DefaultValue;
                }
                else
                {
                    args[i + offset] = arguments[i].As(parameter.ParameterType, context);
                }
            }

            for (var i = n + offset; i < max; i++)
            {
                if (parameters[i].IsOptional)
                {
                    args[i] = parameters[i].DefaultValue;
                }
                else
                {
                    args[i] = parameters[i].ParameterType.GetDefaultValue();
                }
            }

            if (max != parameters.Length)
            {
                var array = Array.CreateInstance(parameters[max].ParameterType.GetElementType(), Math.Max(0, arguments.Length - max));

                for (var i = max; i < arguments.Length; i++)
                {
                    array.SetValue(arguments[i].FromJsValue(), i - max);
                }

                args[max] = array;
            }

            return args;
        }

        private static JsValue[] ExpandInitDict(JsValue[] arguments, ParameterDescription[] parameters, DomInitDictAttribute initDict, Int32 max, Int32 offset)
        {
            var newArgs = new JsValue[max - offset];
            var end = initDict.Offset - offset;
            var obj = arguments[end].AsObject();

            for (var i = 0; i < end; i++)
            {
                newArgs[i] = arguments[i];
            }

            for (var i = end + offset; i < max; i++)
            {
                var p = parameters[i];
                var name = p.Name;

                if (obj.HasProperty(name))
                {
                    newArgs[i - offset] = obj.Get(name);
                }
                else
                {
                    newArgs[i - offset] = JsValue.Undefined;
                }
            }

            arguments = newArgs;
            return arguments;
        }

        public static void AddConstructors(this EngineInstance engine, ObjectInstance ctx, Assembly assembly)
        {
            foreach (var exportedType in assembly.ExportedTypes)
            {
                var enumDefinition = exportedType.GetEnumLiteralDefinition();

                if (enumDefinition != null)
                {
                    SetEnumLiteral(engine, ctx, enumDefinition);
                }
            }

            var definitions = SelectConstructors(assembly.ExportedTypes);

            foreach (var definition in definitions)
            {
                ctx.FastSetProperty(definition.Key, CreateConstructorProperty(engine, definition.Value));
            }
        }

        private static void SetEnumLiteral(EngineInstance engine, ObjectInstance obj, EnumLiteralDefinition enumDefinition)
        {
            var target = obj.Get(enumDefinition.Name) as ObjectInstance;

            if (target == null)
            {
                target = engine.Jint.Intrinsics.Object.Construct(Array.Empty<JsValue>(), JsValue.Undefined);
                obj.FastSetProperty(enumDefinition.Name, new PropertyDescriptor(target, false, true, false));
            }

            foreach (var member in enumDefinition.Members)
            {
                var constant = JsNumber.Create(Convert.ToDouble(member.Value));

                target.FastSetProperty(member.Name, new PropertyDescriptor(
                    constant,
                    false,
                    true,
                    false));
            }
        }

        internal static IReadOnlyDictionary<String, ConstructorDefinition> SelectConstructors(IEnumerable<Type> exportedTypes)
        {
            var selected = new Dictionary<String, ConstructorCandidate>(StringComparer.Ordinal);

            foreach (var exportedType in exportedTypes)
            {
                var definition = exportedType.GetConstructorDefinition();

                if (definition == null)
                {
                    continue;
                }

                var rank = GetExposureRank(exportedType);

                foreach (var name in definition.Names)
                {
                    if (!selected.TryGetValue(name, out var existing) || rank > existing.Rank)
                    {
                        selected[name] = new ConstructorCandidate(definition, rank);
                    }
                }
            }

            var result = new Dictionary<String, ConstructorDefinition>(selected.Count, StringComparer.Ordinal);

            foreach (var selectedName in selected)
            {
                result[selectedName.Key] = selectedName.Value.Definition;
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

        public static void AddConstructorFunctions(this EngineInstance engine, ObjectInstance ctx, Assembly assembly)
        {
            foreach (var exportedType in assembly.ExportedTypes)
            {
                engine.AddConstructorFunction(ctx, exportedType);
            }
        }

        public static void AddInstances(this EngineInstance engine, ObjectInstance obj, Assembly assembly)
        {
            foreach (var exportedType in assembly.ExportedTypes)
            {
                engine.AddInstance(obj, exportedType);
            }
        }

        public static void AddConstructor(this EngineInstance engine, ObjectInstance obj, Type type)
        {
            var enumDefinition = type.GetEnumLiteralDefinition();

            if (enumDefinition != null)
            {
                SetEnumLiteral(engine, obj, enumDefinition);

                return;
            }

            var definition = type.GetConstructorDefinition();

            if (definition != null)
            {
                foreach (var name in definition.Names)
                {
                    obj.FastSetProperty(name, CreateConstructorProperty(engine, definition));
                }
            }
        }

        /// <summary>
        /// The property an exposed type is published under, on the window and on the global object.
        /// A document names a handful of the types an assembly exposes, but a property is registered
        /// for every one of them, so the constructor object behind it is only built once script reads
        /// the property.
        /// </summary>
        /// <remarks>
        /// The attributes are the ones an eagerly written constructor had: enumerable, but neither
        /// writable nor configurable. A lazy descriptor rather than a hand-written custom-valued one
        /// because it stops being lazy once it holds its value, and the engine's global-identifier
        /// cache declines a descriptor that could still compute - permanently, since it has no way to
        /// learn that a custom value became a constant. A type name is exactly the sort of global a
        /// script reads over and over.
        /// </remarks>
        private static PropertyDescriptor CreateConstructorProperty(EngineInstance engine, ConstructorDefinition definition) =>
            PropertyDescriptor.CreateLazy(
                new ConstructorRequest(engine, definition),
                static request => request.Instance.GetDomConstructor(request.Definition),
                PropertyFlag.OnlyEnumerable);

        //  Handed to the factory instead of captured by it, so that the delegate above is the same
        //  one for every type rather than a closure allocated per registered name.
        private readonly struct ConstructorRequest
        {
            public ConstructorRequest(EngineInstance instance, ConstructorDefinition definition)
            {
                Instance = instance;
                Definition = definition;
            }

            public EngineInstance Instance { get; }

            public ConstructorDefinition Definition { get; }
        }

        private readonly struct ConstructorCandidate
        {
            public ConstructorCandidate(ConstructorDefinition definition, Int32 rank)
            {
                Definition = definition;
                Rank = rank;
            }

            public ConstructorDefinition Definition { get; }

            public Int32 Rank { get; }
        }

        public static void AddConstructorFunction(this EngineInstance engine, ObjectInstance obj, Type type)
        {
            var apply = type.GetConstructorFunctionAction();
            apply.Invoke(engine, obj);
        }

        public static void AddInstance(this EngineInstance engine, ObjectInstance obj, Type type)
        {
            var apply = type.GetInstanceAction();
            apply.Invoke(engine, obj);
        }

        /// <summary>
        /// Gets the engine a value belongs to, or null when nothing about it says.
        /// </summary>
        /// <remarks>
        /// A member declared on a shared prototype layout runs on behalf of whichever engine
        /// instantiated it, so it has to work that out from the receiver. A DOM object answers
        /// directly; anything else - the global object, or a plain object given a DOM prototype
        /// by Object.create - answers through the first prototype in its chain that is one of
        /// ours. Only a call made with no receiver at all leaves the question unanswerable.
        /// </remarks>
        public static EngineInstance GetEngineInstance(this JsValue value)
        {
            if (value is IDomProxy proxy)
            {
                return proxy.Instance;
            }

            for (var obj = value as ObjectInstance; obj != null; obj = obj.Prototype)
            {
                var state = DomPrototypeState.Of(obj);

                if (state != null)
                {
                    return state.Instance;
                }
            }

            return null;
        }

        /// <summary>
        /// Invokes a DOM member on behalf of the engine the receiver belongs to.
        /// </summary>
        public static JsValue CallShared(MethodInfo method, JsValue thisObject, JsValue[] arguments)
        {
            var instance = thisObject.GetEngineInstance();

            if (instance == null)
            {
                //  A DOM member torn off its object and called with no receiver - what a browser
                //  answers with a TypeError, and what the engine-bound member this replaces used
                //  to answer by invoking against the window and failing further in.
                throw new JavaScriptException("Illegal invocation.");
            }

            return instance.Call(method, thisObject, arguments);
        }

        public static JsValue Call(this EngineInstance instance, MethodInfo method, JsValue thisObject, JsValue[] arguments)
        {
            if (method != null)
            {
                IDomProxy nodeInstance;

                if (thisObject.Type == Types.Object && thisObject.AsObject() is IDomProxy node)
                {
                    nodeInstance = node;
                }
                else
                {
                    nodeInstance = instance.Window;
                }

                try
                {
                    if (method.IsStatic)
                    {
                        var newArgs = new JsValue[arguments.Length + 1];
                        newArgs[0] = (JsValue)nodeInstance;
                        Array.Copy(arguments, 0, newArgs, 1, arguments.Length);
                        var parameters = instance.BuildArgs(method, newArgs);
                        return method.Invoke(null, parameters).ToJsValue(instance);
                    }
                    else
                    {
                        var parameters = instance.BuildArgs(method, arguments);
                        return method.Invoke(nodeInstance.Value, parameters).ToJsValue(instance);
                    }
                }
                catch (TargetInvocationException)
                {
                    throw new JavaScriptException(instance.Jint.Intrinsics.Error);
                }
            }

            return JsValue.Undefined;
        }
    }
}
