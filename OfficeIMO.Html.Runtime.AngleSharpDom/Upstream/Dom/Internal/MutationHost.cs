namespace AngleSharp.Dom
{
    using AngleSharp.Browser;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.CompilerServices;

    /// <summary>
    /// Couples the mutation events to mutation observers and the event loop.
    /// </summary>
    sealed class MutationHost
    {
        #region Fields

        private static readonly ConditionalWeakTable<IEventLoop, NotificationState> SharedStates = new();
        private readonly NotificationState _state;

        #endregion

        #region ctor

        public MutationHost(IEventLoop loop)
        {
            _state = loop is IMutationMicrotaskScheduler
                ? SharedStates.GetValue(loop, value => new NotificationState(value))
                : new NotificationState(loop);
        }

        #endregion

        #region Properties

        public IEnumerable<MutationObserver> Observers => _state.Observers;

        internal IEnumerable<(MutationObserver Observer, MutationObserver.MutationObserving Registration)> Registrations(INode node) =>
            _state.Observers.SelectMany(observer => observer.ResolveRegistrations(node)
                .Select(entry => (Observer: observer, entry.Registration, entry.Order)))
                .OrderBy(entry => entry.Order)
                .Select(entry => (entry.Observer, entry.Registration));

        #endregion

        #region Methods

        public void Register(MutationObserver observer)
        {
            if (!_state.Observers.Contains(observer))
            {
                _state.Observers.Add(observer);
            }
        }

        public void Unregister(MutationObserver observer)
        {
            if (_state.Observers.Contains(observer))
            {
                _state.Observers.Remove(observer);
            }
        }

        public void MarkPending(MutationObserver observer) => _state.MarkPending(observer);

        public void ScheduleCallback() => _state.Schedule();

        private sealed class NotificationState(IEventLoop loop)
        {
            internal readonly List<MutationObserver> Observers = [];
            private readonly List<MutationObserver> _pending = [];
            private Boolean _queued;

            internal void MarkPending(MutationObserver observer)
            {
                if (!_pending.Contains(observer)) _pending.Add(observer);
            }

            internal void Schedule()
            {
                if (_queued) return;
                if (_pending.Count == 0 && !Observers.Exists(observer => observer.HasTransientRegistrations)) return;
                _queued = true;
                if (loop is IMutationMicrotaskScheduler scheduler) scheduler.EnqueueMutationMicrotask(Dispatch);
                else loop.Enqueue(Dispatch, TaskPriority.Microtask);
            }

            private void Dispatch()
            {
                var pending = _pending.ToArray();
                _pending.Clear();
                _queued = false;
                // Attribute-only subtree observers can have transient registrations
                // even when the removal did not produce an interested record.
                foreach (var observer in Observers.ToArray())
                    if (Array.IndexOf(pending, observer) < 0) observer.ClearTransients();
                foreach (var observer in pending) observer.Trigger();
            }
        }

        #endregion
    }
}
