namespace AngleSharp.Js
{
    using AngleSharp.Js.Cache;
    using Jint;
    using Jint.Native.Object;
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Reflection;

    sealed class PrototypeCache
    {
        private readonly ConcurrentDictionary<Type, ObjectInstance> _prototypes;
        private readonly ConcurrentDictionary<Type, Type> _canonicalTypes;
        private readonly LibrarySet _libs;

        public PrototypeCache(Engine engine, LibrarySet libs)
        {
            _prototypes = new ConcurrentDictionary<Type, ObjectInstance>
            {
                [typeof(Object)] = engine.Intrinsics.Object.PrototypeObject,
            };
            _canonicalTypes = new ConcurrentDictionary<Type, Type>();
            _libs = libs;
        }

        public ObjectInstance GetOrCreate(Type type, Func<Type, ObjectInstance> creator) =>
            _prototypes.GetOrAdd(Canonicalize(type), creator.Invoke);

        //  Memoized per engine rather than globally: the set of libraries a document uses is what
        //  decides the outcome, and that is fixed for an engine but not for the process.
        public Type Canonicalize(Type type) =>
            _canonicalTypes.GetOrAdd(type, m => m.GetDomPrototypeType(_libs.Assemblies));
    }
}
