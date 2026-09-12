namespace AngleSharp.Js
{
    using AngleSharp.Dom;
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Runtime.Descriptors;
    using System;
    using System.Collections.Generic;

    sealed class DomNodeInstance : ObjectInstance, IDomProxy
    {
        private readonly EngineInstance _instance;
        private readonly Object _value;

        private Dictionary<DomEventDefinition, DomEventDefinition.Registration> _eventHandlers;
        private DomPrototypeState _state;

        public DomNodeInstance(EngineInstance engine, Object value)
            : base(engine.Jint)
        {
            _instance = engine;
            _value = value;

            var prototype = engine.GetDomPrototype(value.GetType());
            Prototype = prototype;
            _state = DomPrototypeState.Of(prototype);
        }

        /// <summary>
        /// Gets what this node's prototype knows, or null if a script has given it a prototype
        /// that is not one of ours.
        /// </summary>
        /// <remarks>
        /// Every property lookup on every node asks this - it is what decides whether an indexer
        /// claims the name - so the answer is remembered rather than reached through two type
        /// tests per lookup. The prototype of a node is not fixed, though: a script may hand it
        /// another one, so what is remembered is checked against the prototype in force.
        /// </remarks>
        private DomPrototypeState State
        {
            get
            {
                var state = _state;
                var prototype = Prototype;

                if (state == null || !ReferenceEquals(state.Prototype, prototype))
                {
                    state = DomPrototypeState.Of(prototype);
                    _state = state;
                }

                return state;
            }
        }

        public Object Value => _value;

        public EngineInstance Instance => _instance;

        public override object ToObject() => _value;

        /// <summary>
        /// Gets the handler assigned to this node for the given event, if any.
        /// The handler is per node - the <see cref="DomEventDefinition"/> itself is
        /// shared by every node of the type, in every engine.
        /// </summary>
        public DomEventDefinition.Registration GetEventHandler(DomEventDefinition ev)
        {
            if (_eventHandlers != null && _eventHandlers.TryGetValue(ev, out var registration))
            {
                return registration;
            }

            return null;
        }

        /// <summary>
        /// Assigns the handler for the given event to this node.
        /// </summary>
        public void SetEventHandler(DomEventDefinition ev, DomEventDefinition.Registration registration)
        {
            _eventHandlers = _eventHandlers ?? new Dictionary<DomEventDefinition, DomEventDefinition.Registration>();
            _eventHandlers[ev] = registration;
        }

        /// <summary>
        /// Removes and returns the handler assigned to this node for the given event, if any.
        /// </summary>
        public DomEventDefinition.Registration RemoveEventHandler(DomEventDefinition ev)
        {
            if (_eventHandlers != null && _eventHandlers.TryGetValue(ev, out var registration))
            {
                _eventHandlers.Remove(ev);
                return registration;
            }

            return null;
        }

        public override PropertyDescriptor GetOwnProperty(JsValue property)
        {
            //  An indexer is the only thing that can turn into an own property of the node
            //  itself. The members of the DOM interface live on the prototype, so finding
            //  them is the engine's job - answering them here would make the node claim
            //  every inherited member as its own.
            switch (LookupIndex(property, out var indexed))
            {
                case DomIndexers.IndexerResult.Value:
                    return new PropertyDescriptor(indexed.ToJsValue(_instance), false, false, false);
                case DomIndexers.IndexerResult.Absent:
                    return PropertyDescriptor.Undefined;
                case DomIndexers.IndexerResult.Named:
                    return State.CreateNamedDescriptor(_value, property.ToString());
            }

            return base.GetOwnProperty(property);
        }

        /// <summary>
        /// Hands the engine the value of an own property directly, with no descriptor in
        /// between. Jint asks this wherever it only wants the value, which is every read a
        /// script performs, so an indexed read no longer allocates a descriptor for the
        /// engine to unwrap and drop.
        /// </summary>
        /// <remarks>
        /// Returning false is a statement that the node has no own property of that name,
        /// not that the value was awkward to produce: the engine trusts it, does not ask
        /// again, and continues the read on the prototype. It is exactly what the discarded
        /// descriptor used to prove, which is what lets the engine keep its prototype cache
        /// while a node's own property set lives outside the engine entirely.
        /// </remarks>
        protected override Boolean TryGetOwnPropertyValue(JsValue property, JsValue receiver, out JsValue value)
        {
            switch (LookupIndex(property, out var indexed))
            {
                case DomIndexers.IndexerResult.Value:
                    value = indexed.ToJsValue(_instance);
                    return true;
                case DomIndexers.IndexerResult.Absent:
                    value = JsValue.Undefined;
                    return false;
                case DomIndexers.IndexerResult.Named:
                    //  What the accessor pair's getter would have returned, without building
                    //  the pair. Null means the name is only writable, and reads as undefined.
                    value = indexed == null ? JsValue.Undefined : indexed.ToJsValue(_instance);
                    return true;
            }

            //  What is left is what script or the window mirror wrote onto the node. Jint's
            //  own fallback would route back through GetOwnProperty and ask the indexers a
            //  second time, so the property bag is read here instead.
            var descriptor = base.GetOwnProperty(property);

            if (ReferenceEquals(descriptor, PropertyDescriptor.Undefined))
            {
                value = JsValue.Undefined;
                return false;
            }

            if (descriptor.Get == null)
            {
                //  Value also covers a descriptor that produces its value on read - the
                //  deferred constructors on the window are of that kind.
                value = descriptor.Value ?? JsValue.Undefined;
                return true;
            }

            //  An accessor has to be invoked against the receiver, and that is the engine's
            //  job. Script installs those on nodes - React tracks input values that way.
            return base.TryGetOwnPropertyValue(property, receiver, out value);
        }

        /// <summary>
        /// Answers whether an own property exists, and whether it enumerates, without
        /// building a descriptor for it. Backs "in", hasOwnProperty, Object.assign, object
        /// spread and JSON.stringify.
        /// </summary>
        /// <remarks>
        /// The answer has to be the one <see cref="GetOwnProperty"/> would give at the same
        /// moment - the engine does not verify it, and a wrong "missing" silently drops the
        /// property from every one of those. Both go through the same lookup for that reason.
        /// </remarks>
        protected override OwnPropertyProbe ProbeOwnProperty(JsValue property)
        {
            switch (LookupIndex(property, out _))
            {
                case DomIndexers.IndexerResult.Value:
                    //  The attributes GetOwnProperty gives an indexed entry.
                    return OwnPropertyProbe.NonEnumerable;
                case DomIndexers.IndexerResult.Absent:
                    return OwnPropertyProbe.Missing;
                case DomIndexers.IndexerResult.Named:
                    return OwnPropertyProbe.NonEnumerable;
            }

            var descriptor = base.GetOwnProperty(property);

            if (ReferenceEquals(descriptor, PropertyDescriptor.Undefined))
            {
                return OwnPropertyProbe.Missing;
            }

            return descriptor.Enumerable ? OwnPropertyProbe.Enumerable : OwnPropertyProbe.NonEnumerable;
        }

        private DomIndexers.IndexerResult LookupIndex(JsValue property, out Object indexed)
        {
            var state = State;
            var indexers = state?.Indexers;

            if (indexers != null)
            {
                return indexers.TryGetFromIndex(state.Prototype, _value, property, out indexed);
            }

            indexed = null;
            return DomIndexers.IndexerResult.None;
        }

        protected override void SetOwnProperty(JsValue property, PropertyDescriptor desc)
        {
            var state = State;
            var indexers = state?.Indexers;

            if (indexers != null && indexers.TrySetToIndex(state.Prototype, _instance, _value, property, desc.Value))
            {
                return;
            }

            base.SetOwnProperty(property, desc);

            if (_value is IWindow)
            {
                _instance.Jint.Global.FastSetProperty(property, desc);
            }
        }

        public new void FastSetProperty(string name, PropertyDescriptor value)
        {
            base.FastSetProperty(name, value);

            if (_value is IWindow)
            {
                _instance.Jint.Global.FastSetProperty(name, value);
            }
        }

        public new void FastSetProperty(JsValue property, PropertyDescriptor value)
        {
            base.FastSetProperty(property, value);

            if (_value is IWindow)
            {
                _instance.Jint.Global.FastSetProperty(property, value);
            }
        }

        public new void FastSetDataProperty(string name, JsValue value)
        {
            base.FastSetDataProperty(name, value);

            if (_value is IWindow)
            {
                _instance.Jint.Global.FastSetProperty(name, new PropertyDescriptor(value, true, true, true));
            }
        }
    }
}
