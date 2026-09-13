namespace AngleSharp.Dom
{
    using AngleSharp.Attributes;
    using System;
    using System.Collections.Generic;
    using System.Threading;

    /// <summary>
    /// MutationObserver provides developers a way to react to changes in a
    /// DOM.
    /// </summary>
    [DomName("MutationObserver")]
    public sealed class MutationObserver
    {
        #region Fields

        private readonly Queue<IMutationRecord> _records;
        private readonly MutationCallback _callback;
        private readonly List<MutationObserving> _observing;
        private static Int64 _registrationOrder;

        #endregion

        #region ctor

        /// <summary>
        /// Creates a new mutation observer with the provided callback.
        /// </summary>
        /// <param name="callback">The callback to trigger.</param>
        [DomConstructor]
        public MutationObserver(MutationCallback callback)
        {
            _records = new Queue<IMutationRecord>();
            _callback = callback ?? throw new ArgumentNullException(nameof(callback));
            _observing = [];
        }

        #endregion

        #region Properties

        internal Boolean HasTransientRegistrations => _observing.Exists(observing => observing.TransientNodes.Count != 0);


        private MutationObserving? this[INode node]
        {
            get
            {
                foreach (var observing in _observing)
                {
                    if (Object.ReferenceEquals(observing.Target, node))
                    {
                        return observing;
                    }
                }

                return null;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Queues a record.
        /// </summary>
        /// <param name="record">The record to queue up.</param>
        internal void Enqueue(MutationRecord record)
        {
            if (_records.Count > 0)
            {
                //Here we could schedule a callback!
            }

            _records.Enqueue(record);
        }

        /// <summary>
        /// Triggers the execution if the queue is not-empty.
        /// </summary>
        internal void Trigger()
        {
            var records = _records.ToArray();
            _records.Clear();
            ClearTransients();

            if (records.Length != 0)
            {
                _callback(records, this);
            }
        }

        /// <summary>
        /// Gets every direct or source-associated transient registration on a node.
        /// </summary>
        /// <param name="node">The node of interest.</param>
        /// <returns>The source registration and its order on this node.</returns>
        internal IEnumerable<(MutationObserving Registration, Int64 Order)> ResolveRegistrations(INode node)
        {
            foreach (var observing in _observing)
            {
                if (Object.ReferenceEquals(observing.Target, node))
                    yield return (observing, observing.Order);
                if (observing.TransientNodes.TryGetValue(node, out var order))
                    yield return (observing, order);
            }
        }

        /// <summary>
        /// Adds a transient registration while retaining its original source.
        /// </summary>
        /// <param name="source">The direct registration supplying the options.</param>
        /// <param name="node">
        /// The node to observe as a transient observer.
        /// </param>
        internal void AddTransient(MutationObserving source, INode node)
        {
            if (source.Options.IsObservingSubtree && !source.TransientNodes.ContainsKey(node))
                source.TransientNodes.Add(node, Interlocked.Increment(ref _registrationOrder));
        }

        /// <summary>
        /// Clears all transient observers.
        /// </summary>
        internal void ClearTransients()
        {
            foreach (var observing in _observing)
            {
                observing.TransientNodes.Clear();
            }
        }

        /// <summary>
        /// Stops the MutationObserver instance from receiving
        /// notifications of DOM mutations. Until the observe()
        /// method is used again, observer's callback will not be invoked.
        /// </summary>
        [DomName("disconnect")]
        public void Disconnect()
        {
            foreach (var observing in _observing)
            {
                var node = (Node)observing.Target;
                (node as Document ?? node.Owner)?.Mutations.Unregister(this);
            }

            _records.Clear();
            _observing.Clear();
        }

        /// <summary>
        /// Registers the MutationObserver instance to receive notifications of
        /// DOM mutations on the specified node.
        /// </summary>
        /// <param name="target">
        /// The Node on which to observe DOM mutations.
        /// </param>
        /// <param name="childList">
        /// If additions and removals of the target node's child elements
        /// (including text nodes) are to be observed.
        /// </param>
        /// <param name="subtree">
        /// If mutations to not just target, but also target's descendants are
        /// to be observed.
        /// </param>
        /// <param name="attributes">
        /// If mutations to target's attributes are to be observed.
        /// </param>
        /// <param name="characterData">
        /// If mutations to target's data are to be observed.
        /// </param>
        /// <param name="attributeOldValue">
        /// If attributes is set to true and target's attribute value before
        /// the mutation needs to be recorded.
        /// </param>
        /// <param name="characterDataOldValue">
        /// If characterData is set to true and target's data before the
        /// mutation needs to be recorded.
        /// </param>
        /// <param name="attributeFilter">
        /// The attributes to observe. If this is not set, then all attributes
        /// are being observed.
        /// </param>
        [DomName("observe")]
        [DomInitDict(offset: 1)]
        public void Connect(INode target, Boolean childList = false, Boolean subtree = false, Boolean? attributes = null, Boolean? characterData = null, Boolean? attributeOldValue = null, Boolean? characterDataOldValue = null, IEnumerable<String>? attributeFilter = null)
        {
            if (target is Node node)
            {
                var oldCharacterData = characterDataOldValue ?? false;
                var oldAttributeValue = attributeOldValue ?? false;

                var options = new MutationOptions
                {
                    IsObservingChildNodes = childList,
                    IsObservingSubtree = subtree,
                    IsExaminingOldCharacterData = oldCharacterData,
                    IsExaminingOldAttributeValue = oldAttributeValue,
                    IsObservingCharacterData = characterData ?? oldCharacterData,
                    IsObservingAttributes = attributes ?? (oldAttributeValue || attributeFilter != null),
                    AttributeFilters = attributeFilter
                };

                if (options.IsExaminingOldAttributeValue && !options.IsObservingAttributes)
                {
                    throw new DomException(DomError.TypeMismatch);
                }

                if (options.AttributeFilters != null && !options.IsObservingAttributes)
                {
                    throw new DomException(DomError.TypeMismatch);
                }

                if (options.IsExaminingOldCharacterData && !options.IsObservingCharacterData)
                {
                    throw new DomException(DomError.TypeMismatch);
                }

                if (options.IsInvalid)
                {
                    throw new DomException(DomError.Syntax);
                }

                var owner = node as Document ?? node.Owner;
                if (owner is null)
                {
                    throw new DomException(DomError.HierarchyRequest);
                }

                owner.Mutations.Register(this);

                var existing = this[target];

                if (existing != null)
                {
                    existing.TransientNodes.Clear();
                    existing.Options = options;
                }
                else
                    _observing.Add(new MutationObserving(target, options, Interlocked.Increment(ref _registrationOrder)));
            }
        }

        /// <summary>
        /// Empties the MutationObserver instance's record queue and returns
        /// what was in there.
        /// </summary>
        /// <returns>Returns an Array of MutationRecords.</returns>
        [DomName("takeRecords")]
        public IEnumerable<IMutationRecord> Flush()
        {
            while (_records.Count > 0)
            {
                yield return _records.Dequeue();
            }
        }

        #endregion

        #region Options

        internal struct MutationOptions
        {
            public Boolean IsObservingChildNodes;
            public Boolean IsObservingSubtree;
            public Boolean IsObservingCharacterData;
            public Boolean IsObservingAttributes;
            public Boolean IsExaminingOldCharacterData;
            public Boolean IsExaminingOldAttributeValue;
            public IEnumerable<String>? AttributeFilters;

            public readonly Boolean IsInvalid => !IsObservingAttributes && !IsObservingCharacterData && !IsObservingChildNodes;
        }

        internal sealed class MutationObserving(INode target, MutationOptions options, Int64 order)
        {
            public INode Target { get; } = target;
            public MutationOptions Options { get; set; } = options;
            public Int64 Order { get; } = order;
            public Dictionary<INode, Int64> TransientNodes { get; } = [];
        }

        #endregion
    }
}
