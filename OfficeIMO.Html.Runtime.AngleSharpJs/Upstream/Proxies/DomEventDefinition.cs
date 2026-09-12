namespace AngleSharp.Js
{
    using AngleSharp.Dom;
    using Jint;
    using Jint.Native;
    using Jint.Native.Function;
    using System;
    using System.Reflection;

    /// <summary>
    /// One event-handler member of a DOM type - "onclick" and its kin - as the accessor pair
    /// a prototype declares for it.
    /// </summary>
    /// <remarks>
    /// The pair is process-shared along with the rest of the type's members, so this object
    /// must stay free of anything belonging to an engine: it holds the two reflected accessors
    /// and derives everything else from the receiver. It doubles as the key a node files its
    /// handler under, which is what keeps the handler itself - the one thing here that is per
    /// node and per engine - on the node.
    /// </remarks>
    sealed class DomEventDefinition
    {
        private readonly MethodInfo _addHandler;
        private readonly MethodInfo _removeHandler;

        public DomEventDefinition(MethodInfo addHandler, MethodInfo removeHandler)
        {
            _addHandler = addHandler;
            _removeHandler = removeHandler;
        }

        public JsValue GetHandler(JsValue thisObject, JsValue[] arguments)
        {
            var node = thisObject.As<DomNodeInstance>();
            var registration = node?.GetEventHandler(this);
            return registration?.Function ?? JsValue.Null;
        }

        public JsValue SetHandler(JsValue thisObject, JsValue[] arguments)
        {
            var node = thisObject.As<DomNodeInstance>();
            var value = arguments.Length > 0 ? arguments[0] : JsValue.Undefined;

            if (node != null)
            {
                var previous = node.RemoveEventHandler(this);

                if (previous != null)
                {
                    _removeHandler?.Invoke(node.Value, new Object[] { previous.Handler });
                }

                if (value is Function function)
                {
                    var engine = node.Instance;

                    DomEventHandler handler = (s, ev) =>
                    {
                        var sender = s.ToJsValue(engine);
                        var args = ev.ToJsValue(engine);
                        function.Call(sender, new[] { args });
                    };

                    node.SetEventHandler(this, new Registration(function, handler));
                    _addHandler?.Invoke(node.Value, new Object[] { handler });
                }
            }

            return value;
        }

        /// <summary>
        /// The handler currently assigned to a single node for a single event.
        /// </summary>
        public sealed class Registration
        {
            public Registration(Function function, DomEventHandler handler)
            {
                Function = function;
                Handler = handler;
            }

            /// <summary>
            /// The function that was assigned, as it has to be handed back on read.
            /// </summary>
            public Function Function { get; }

            /// <summary>
            /// The listener that was subscribed, as it has to be handed back on removal.
            /// </summary>
            public DomEventHandler Handler { get; }
        }
    }
}
