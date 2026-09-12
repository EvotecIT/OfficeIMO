namespace AngleSharp.Js
{
    using AngleSharp.Js.Cache;
    using Jint;
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Native.Symbol;
    using Jint.Runtime.Descriptors;
    using System;
    using System.Reflection;

    /// <summary>
    /// The JS object standing for a DOM collection - a node list, an HTML collection, a token
    /// list, a named node map. Jint derives the whole array-like property model from the two
    /// members below, so indices and "length" are answered out of the collection itself with no
    /// descriptor, no reflection over the prototype and no exception at the end.
    /// </summary>
    /// <remarks>
    /// It is deliberately not an Array, which is what a browser reports too: "Array.isArray" is
    /// false and the prototype chain is still the DOM one. What it does gain is the iterator, so
    /// "for...of", spread and "Array.from" work on a collection the way they do in a browser.
    /// </remarks>
    sealed class DomCollectionInstance : ArrayLikeObject, IDomProxy
    {
        private readonly EngineInstance _instance;
        private readonly Object _value;
        private readonly IndexedCollection _collection;

        private DomPrototypeState _state;

        public DomCollectionInstance(EngineInstance engine, Object value, IndexedCollection collection)
            : base(engine.Jint)
        {
            _instance = engine;
            _value = value;
            _collection = collection;

            var prototype = engine.GetDomPrototype(value.GetType());
            Prototype = prototype;
            _state = DomPrototypeState.Of(prototype);

            //  A DOM collection is iterable in a browser, and wiring the array iterator is what
            //  the specifications themselves prescribe for it. Jint recognises that exact
            //  function and takes its array-like path through TryGetIndex for it.
            FastSetProperty(GlobalSymbolRegistry.Iterator, new PropertyDescriptor(
                engine.Jint.Intrinsics.Array.PrototypeObject.Get(GlobalSymbolRegistry.Iterator),
                true, false, true));
        }

        public Object Value => _value;

        public EngineInstance Instance => _instance;

        //  Remembered rather than reached through two type tests per lookup, and checked against
        //  the prototype in force because a script may hand the collection another one.
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

        public override Object ToObject() => _value;

        public override UInt32 Length
        {
            get
            {
                var length = _collection.GetLength(_value);
                return length > 0 ? (UInt32)length : 0;
            }
        }

        public override Boolean TryGetIndex(UInt32 index, out JsValue value)
        {
            //  Length is read live, so an index can go out of range between the two calls, and
            //  a collection whose accessor is stricter than its length still has to answer
            //  rather than throw.
            if (index < Length)
            {
                try
                {
                    value = _collection.GetItem(_value, (Int32)index).ToJsValue(_instance);
                    return true;
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }

            value = JsValue.Undefined;
            return false;
        }

        /// <summary>
        /// Whether the collection currently has an element at the index. Containment in an
        /// indexed collection is the length comparison and nothing else, so the questions that
        /// only want a yes or a no - "in", hasOwnProperty, the key enumerations, the hole test
        /// the Array.prototype generics run per element - stop projecting an element, wrapping
        /// it in a JS value and looking it up in the identity cache just to drop it again.
        /// </summary>
        /// <remarks>
        /// The one case where this could disagree with <see cref="TryGetIndex"/> is a collection
        /// whose accessor refuses an index its own length admits. No DOM collection does that,
        /// and a Debug build of Jint checks both directions of the agreement on every probe, so
        /// the whole test suite is the proof.
        /// </remarks>
        protected override Boolean HasIndex(UInt32 index) => index < Length;

        /// <summary>
        /// Named entries - "document.forms.login", "el.attributes.title" - are the one thing a
        /// collection answers that the array-like model knows nothing about.
        /// </summary>
        /// <remarks>
        /// Indices and "length" have to reach the base implementation untouched, which is why
        /// the base is asked first: it owns them, and a Debug build of Jint verifies that it was
        /// allowed to answer them.
        /// </remarks>
        public override PropertyDescriptor GetOwnProperty(JsValue property)
        {
            var descriptor = base.GetOwnProperty(property);

            if (!ReferenceEquals(descriptor, PropertyDescriptor.Undefined))
            {
                return descriptor;
            }

            //  The base has answered every array-index key authoritatively - an index past the
            //  end is a miss its value and existence hooks already committed to - so the string
            //  indexer must not be asked about one. WebIDL says the same: an object supporting
            //  indexed properties never serves an array-index name from its named getter. An
            //  element whose id is "5" is findable through a non-index name, never through 5.
            if (IsArrayIndexName(property))
            {
                return PropertyDescriptor.Undefined;
            }

            var state = State;
            var indexers = state?.Indexers;

            if (indexers != null && indexers.TryGetFromNamedIndex(state.Prototype, _value, property, out _))
            {
                return state.CreateNamedDescriptor(_value, property.ToString());
            }

            return PropertyDescriptor.Undefined;
        }

        //  The test the engine's array-like model applies to a key, mirrored exactly so this
        //  class and its base always classify a name the same way: a number holding a UInt32
        //  below the maximum, or its canonical string form - no sign, no space, no leading zero.
        private static Boolean IsArrayIndexName(JsValue property)
        {
            if (property.IsNumber())
            {
                var value = property.AsNumber();
                var index = (UInt32)value;
                return value == index && index != UInt32.MaxValue;
            }

            if (property.IsString())
            {
                var name = property.ToString();

                if (name.Length == 0 || (name.Length > 1 && (name[0] < '1' || name[0] > '9')))
                {
                    return false;
                }

                return UInt32.TryParse(name, out var index) && index != UInt32.MaxValue;
            }

            return false;
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
        }
    }
}
