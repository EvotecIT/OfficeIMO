namespace AngleSharp.Js
{
    using AngleSharp.Js.Cache;
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Runtime.Descriptors;
    using Jint.Runtime.Interop;
    using System;

    /// <summary>
    /// What one engine's prototype for a DOM type carries beyond its members: which engine it
    /// belongs to, which type it stands for, and the constructor object published under the
    /// type's name.
    /// </summary>
    /// <remarks>
    /// The member layout is shared by every engine, so none of this can live on it - all three
    /// are either engine-affine or, in the constructor's case, an object of exactly one engine.
    /// It hangs off the prototype instead, which is what
    /// <see cref="JsObjectShape.SetHostState(ObjectInstance, Object)"/> exists for, so it is
    /// reachable from any receiver by walking up to the prototype that declares the member being
    /// served - the way a shared member implementation finds its engine.
    /// </remarks>
    sealed class DomPrototypeState
    {
        private readonly EngineInstance _instance;
        private readonly ObjectInstance _prototype;
        private readonly DomShape _shape;
        private readonly DomIndexers _indexers;

        private DomConstructorInstance _constructor;

        public DomPrototypeState(EngineInstance instance, ObjectInstance prototype, DomShape shape)
        {
            _instance = instance;
            _prototype = prototype;
            _shape = shape;

            //  Copied out rather than reached through the shape: every property lookup on every
            //  node of this type asks whether an indexer claims the name, and for almost every
            //  type the answer is the null below.
            _indexers = shape.Indexers;
        }

        /// <summary>
        /// Gets the state of a prototype, or null for an object that is not one - which is the
        /// answer a walk up a prototype chain needs when it runs off the end.
        /// </summary>
        public static DomPrototypeState Of(ObjectInstance prototype) =>
            JsObjectShape.GetHostState(prototype) as DomPrototypeState;

        /// <summary>
        /// Gets the engine the prototype belongs to.
        /// </summary>
        public EngineInstance Instance => _instance;

        /// <summary>
        /// Gets the prototype this state belongs to, which is how a proxy holding the state
        /// checks that it is still the state of the prototype it currently inherits from.
        /// </summary>
        public ObjectInstance Prototype => _prototype;

        /// <summary>
        /// Gets what the prototype is built from, shared with every other engine.
        /// </summary>
        public DomShape Shape => _shape;

        /// <summary>
        /// Gets the indexers of the type, or null if it has none.
        /// </summary>
        public DomIndexers Indexers => _indexers;

        /// <summary>
        /// Builds the descriptor for a name the string indexer claims. It is an accessor pair
        /// rather than a value so that both directions stay live: the getter re-reads the
        /// indexer, and the setter is what forwards a write to it.
        /// </summary>
        /// <remarks>
        /// It lives here rather than on the indexers because the pair needs an engine as well as
        /// the indexers, and this object is exactly the two of them together - so the closure
        /// behind the pair carries one reference for both, on a path that builds it per read.
        /// </remarks>
        public PropertyDescriptor CreateNamedDescriptor(Object value, String index)
        {
            var args = new Object[] { index };

            //  With nothing to forward a write to there is nothing an accessor pair would add:
            //  the value is read at the same moment either way, and the pair costs two function
            //  objects on a path that runs per read.
            if (!_indexers.HasNamedSetter)
            {
                var current = _indexers.ReadNamed(value, args);
                return new PropertyDescriptor(
                    current == null ? JsValue.Undefined : current.ToJsValue(_instance), false, false, false);
            }

            var getter = new ClrFunction(_instance.Jint, index, (obj, values) =>
            {
                var current = _indexers.ReadNamed(value, args);
                return current == null ? JsValue.Undefined : current.ToJsValue(_instance);
            });

            var setter = new ClrFunction(_instance.Jint, index, (obj, values) =>
            {
                var valueToSet = values.Length > 0 ? values[0] : JsValue.Undefined;
                _indexers.WriteNamed(value, new Object[] { index, valueToSet }, _instance);
                return valueToSet;
            });

            return new GetSetPropertyDescriptor(getter, setter, false, false);
        }

        /// <summary>
        /// Gets the constructor object of the type this prototype belongs to, building it on
        /// first ask. Holding it here is what keeps the one script reads off the window and the
        /// one an instance reports as its "constructor" the same object.
        /// </summary>
        public DomConstructorInstance GetConstructor(ConstructorDefinition definition)
        {
            if (_constructor == null)
            {
                _constructor = new DomConstructorInstance(_instance, definition);

                //  A type the engine's libraries expose under no name of their own gets no
                //  "constructor" slot in the shape, so naming its constructor is the moment the
                //  property has to appear - which is what the constructor object used to do for
                //  every type. A shaped object takes an undeclared name without losing its shape.
                if (_shape.Constructor == null)
                {
                    _prototype.FastSetProperty("constructor", new PropertyDescriptor(_constructor, true, false, true));
                }
            }

            return _constructor;
        }
    }
}
