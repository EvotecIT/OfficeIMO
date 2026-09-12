namespace AngleSharp.Js
{
    using Jint.Native;
    using System;
    using System.Collections.Concurrent;
    using System.Reflection;

    /// <summary>
    /// One of the indexers a prototype offers, invoked against whichever object is being
    /// indexed rather than against the type the prototype was created for.
    /// </summary>
    /// <remarks>
    /// A prototype stands for a DOM type, not for a single class - col and colgroup are
    /// both an HTMLTableColElement, and every element class carrying no name of its own
    /// shares the prototype of its nearest named ancestor. An accessor bound to one of
    /// those classes cannot be invoked on the others, so the target decides.
    /// </remarks>
    sealed class Indexer
    {
        private readonly MethodInfo _declared;
        private readonly ConcurrentDictionary<Type, MethodInfo> _resolved;

        public Indexer(MethodInfo declared)
        {
            _declared = declared;

            //  A public accessor is dispatched virtually and works on any implementation;
            //  only an explicitly re-implemented one has to be looked up per target.
            _resolved = declared.IsPublic ? null : new ConcurrentDictionary<Type, MethodInfo>();
        }

        public Object Invoke(Object target, Object[] arguments, EngineInstance engine = null)
        {
            var method = Resolve(target.GetType());

            if (engine != null)
            {
                var parameters = method.GetParameters();
                var converted = new Object[arguments.Length];

                for (var i = 0; i < arguments.Length; i++)
                {
                    if (arguments[i] is JsValue value)
                    {
                        converted[i] = value.As(parameters[i].ParameterType, engine);
                    }
                    else
                    {
                        converted[i] = arguments[i];
                    }
                }

                arguments = converted;
            }

            return method.Invoke(target, arguments);
        }

        private MethodInfo Resolve(Type targetType) =>
            _resolved == null ? _declared : _resolved.GetOrAdd(targetType, ResolveCore);

        private MethodInfo ResolveCore(Type targetType)
        {
            //  An interface may re-implement a member of one of its own base interfaces
            //  explicitly, e.g. "T IReadOnlyList<T>.this[Int32 index]" declared on an
            //  IHtmlCollection<T>. Such a member is private and abstract - invoking it
            //  reflectively throws an EntryPointNotFoundException because the actual
            //  implementation lives in a different slot.
            var name = _declared.Name;
            var simpleName = name.Substring(name.LastIndexOf('.') + 1);
            var parameters = _declared.GetParameters();
            var parameterTypes = new Type[parameters.Length];

            for (var i = 0; i < parameters.Length; i++)
            {
                parameterTypes[i] = parameters[i].ParameterType;
            }

            return targetType.GetRuntimeMethod(simpleName, parameterTypes) ?? _declared;
        }
    }
}
