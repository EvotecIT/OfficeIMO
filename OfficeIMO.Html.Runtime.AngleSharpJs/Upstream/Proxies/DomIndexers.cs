namespace AngleSharp.Js
{
    using Jint.Native;
    using Jint.Native.Object;
    using Jint.Native.Symbol;
    using Jint.Runtime.Descriptors;
    using Jint.Runtime.Interop;
    using System;
    using System.Reflection;

    /// <summary>
    /// The indexers a DOM type offers - the numeric one a collection is read through, the
    /// string one a named entry is found by - resolved once for the type.
    /// </summary>
    /// <remarks>
    /// Everything here is derived from the type's accessors and from nothing else, so an
    /// instance of this class describes the type rather than any one engine, and is shared by
    /// every prototype built from that type. Only the objects it produces - a value, a
    /// descriptor - belong to the engine that asked for them, which is why the engine is a
    /// parameter of the methods that produce one and of no others.
    /// </remarks>
    sealed class DomIndexers
    {
        private readonly Indexer _numeric;
        private readonly Indexer _namedGetter;
        private readonly Indexer _namedSetter;

        public DomIndexers(Indexer numeric, Indexer namedGetter, Indexer namedSetter)
        {
            _numeric = numeric;
            _namedGetter = namedGetter;
            _namedSetter = namedSetter;
        }

        /// <summary>
        /// Asks the indexers about one property name, against the object being indexed rather
        /// than against the prototype declaring them.
        /// </summary>
        /// <remarks>
        /// The raw CLR value comes back instead of a <see cref="PropertyDescriptor"/> because
        /// only one of the three callers wants a descriptor: the engine asks a host object for
        /// the value of an own property and for its mere existence separately, and building a
        /// descriptor for either is pure waste - and for a named entry it builds two
        /// <see cref="ClrFunction"/>s with it.
        /// </remarks>
        public IndexerResult TryGetFromIndex(ObjectInstance prototype, Object value, JsValue property, out Object result)
        {
            result = null;

            //  Reached for every property lookup on every node of a type that has any indexer at
            //  all, and a write-only string indexer answers no read. Nothing below may run before
            //  that is ruled out - not even turning the key into a String.
            if (_numeric == null && _namedGetter == null)
            {
                return IndexerResult.None;
            }

            //  A symbol is never an index, and JsSymbol.ToString() composes "Symbol(...)"
            //  every time it is asked. Library code probes Symbol.toStringTag constantly.
            if (property is JsSymbol)
            {
                return IndexerResult.None;
            }

            var index = property.ToString();

            //  If we have a numeric indexer and the property is numeric
            if (_numeric != null && Int32.TryParse(index, out var numericIndex))
            {
                try
                {
                    var args = new Object[] { numericIndex };
                    result = _numeric.Invoke(value, args);
                    return IndexerResult.Value;
                }
                catch (TargetInvocationException ex)
                {
                    if (ex.InnerException is ArgumentOutOfRangeException)
                    {
                        return IndexerResult.Absent;
                    }

                    throw;
                }
            }

            //  Else a string property
            return TryGetFromNamedIndex(prototype, value, property, out result) ? IndexerResult.Named : IndexerResult.None;
        }

        /// <summary>
        /// The string-indexer half on its own, for a collection whose numeric half the engine
        /// owns and must not be asked for twice.
        /// </summary>
        public Boolean TryGetFromNamedIndex(ObjectInstance prototype, Object value, JsValue property, out Object result)
        {
            result = null;

            if (_namedGetter == null || property is JsSymbol)
            {
                return false;
            }

            var index = property.ToString();

            //  If we have a string indexer and no property exists for this name then use the string indexer
            //  Jint possibly has a limitation here - if an object has a string indexer.  How do we know whether to use the defined indexer or a property?
            //  Eg. object.callMethod1()  vs  object['callMethod1'] is not necessarily the same if the object has a string indexer?? (I'm not an ECMA expert!)
            //  node.attributes is one such object - has both a string and numeric indexer
            //  This GetOwnProperty override might need an additional parameter to let us know this was called via an indexer
            //
            //  HasProperty takes a JsValue, so handing it the String would build one per
            //  lookup - and every read of an ordinary member of an indexed collection, say
            //  the length a loop tests, comes through here. The engine already passed one in.
            //  A shaped prototype answers it off the member layout, without materialising the
            //  member the answer is about.
            if (prototype.HasProperty(property as JsString ?? (JsValue)index))
            {
                return false;
            }

            var valueAtIndex = _namedGetter.Invoke(value, new Object[] { index });

            if (valueAtIndex == null && _namedSetter == null)
            {
                return false;
            }

            //  Null with a setter present is still an own property - one that reads as
            //  undefined - because a write to it still has to reach the setter.
            result = valueAtIndex;
            return true;
        }

        /// <summary>
        /// Whether a write to a name the string indexer claims has anywhere to go.
        /// </summary>
        public Boolean HasNamedSetter => _namedSetter != null;

        /// <summary>
        /// Reads one name through the string indexer. Called by the descriptor
        /// <see cref="DomPrototypeState.CreateNamedDescriptor"/> hands out, which keeps its
        /// getter live rather than freezing the value it read.
        /// </summary>
        public Object ReadNamed(Object value, Object[] args) => _namedGetter.Invoke(value, args);

        /// <summary>
        /// Writes one name through the string indexer.
        /// </summary>
        public void WriteNamed(Object value, Object[] args, EngineInstance engine) =>
            _namedSetter.Invoke(value, args, engine);

        public Boolean TrySetToIndex(ObjectInstance prototype, EngineInstance engine, Object value, JsValue property, JsValue newValue)
        {
            if (_namedSetter == null)
            {
                return false;
            }

            var index = property.ToString();

            if (prototype.HasProperty(property as JsString ?? (JsValue)index))
            {
                return false;
            }

            _namedSetter.Invoke(value, new Object[] { index, newValue }, engine);
            return true;
        }

        /// <summary>
        /// What the indexers of a prototype have to say about a property name.
        /// </summary>
        internal enum IndexerResult
        {
            /// <summary>
            /// No indexer claims the name, so the ordinary own-property lookup decides.
            /// </summary>
            None,
            /// <summary>
            /// The numeric indexer answered, and the answer is the value handed back.
            /// </summary>
            Value,
            /// <summary>
            /// The numeric indexer claims the name but has nothing at it, so the object has no
            /// own property of that name and the ordinary lookup is not consulted either.
            /// </summary>
            Absent,
            /// <summary>
            /// The string indexer claims the name. The value handed back is what it reads as
            /// right now, and may be null - a name the indexer only accepts writes for is
            /// still an own property, described by an accessor pair.
            /// </summary>
            Named
        }
    }
}
