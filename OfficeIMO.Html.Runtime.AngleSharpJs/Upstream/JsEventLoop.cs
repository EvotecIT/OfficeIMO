namespace AngleSharp.Js
{
    using AngleSharp.Browser;
    using AngleSharp.Common;
    using System;
    using System.Collections.Generic;
    using System.Threading;

    /// <summary>
    /// A thread-based event loop implementation.
    /// </summary>
    public sealed class JsEventLoop : IEventLoop, IDisposable
    {
        //  Scripts run on this thread, and the JS call stack is the native one. The usual
        //  1 MB holds roughly a thousand JavaScript frames, well short of what a browser
        //  offers, and every frame beyond it costs the engine a hop onto a fresh stack.
        //  The size is reserved address space rather than memory, but a 32 bit process has
        //  little of it to spare when it runs many loops, so only a 64 bit one is enlarged;
        //  zero leaves the thread with the default of the process.
        private static readonly Int32 DefaultMaxStackSize = IntPtr.Size == 8 ? 16 * 1024 * 1024 : 0;

        private readonly Dictionary<TaskPriority, Queue<LoopEntry>> _queues = new Dictionary<TaskPriority, Queue<LoopEntry>>();
        private readonly Object _lockObj = new Object();
        private readonly Action<Exception> _trackError;
        private CancellationTokenSource _cts;

        /// <summary>
        /// Creates a new event loop thread.
        /// </summary>
        public JsEventLoop()
            : this(DefaultMaxStackSize, null)
        {
        }

        /// <summary>
        /// Creates a new event loop thread that reports task errors to the given context.
        /// </summary>
        /// <param name="context">The browsing context to report errors to.</param>
        public JsEventLoop(IBrowsingContext context)
            : this(DefaultMaxStackSize, context == null ? null : new Action<Exception>(context.TrackError))
        {
        }

        /// <summary>
        /// Creates a new event loop thread with the given stack size.
        /// </summary>
        /// <param name="maxStackSize">The stack size of the thread running the scripts.</param>
        public JsEventLoop(Int32 maxStackSize)
            : this(maxStackSize, null)
        {
        }

        private JsEventLoop(Int32 maxStackSize, Action<Exception> trackError)
        {
            _trackError = trackError;
            var thread = new Thread(Runner, maxStackSize)
            {
                IsBackground = true,
                Name = "AngleSharpEventLoop",
#if !NETSTANDARD1_3
                Priority = ThreadPriority.Highest,
#endif
            };
            _cts = new CancellationTokenSource();
            thread.Start(_cts.Token);
        }

        ICancellable IEventLoop.Enqueue(Action<CancellationToken> action, TaskPriority priority)
        {
            var entry = new LoopEntry(action, _trackError);

            lock (_lockObj)
            {
                if (!_queues.TryGetValue(priority, out var entries))
                {
                    entries = new Queue<LoopEntry>();
                    _queues.Add(priority, entries);
                }

                entries.Enqueue(entry);
            }

            return entry;
        }

        private LoopEntry Dequeue()
        {
            LoopEntry Dequeue(TaskPriority priority)
            {
                if (_queues.ContainsKey(priority) && _queues[priority].Count != 0)
                {
                    return _queues[priority].Dequeue();
                }

                return null;
            }

            var current = default(LoopEntry);

            lock (_lockObj)
            {
                current =
                    Dequeue(TaskPriority.Critical) ??
                    Dequeue(TaskPriority.Microtask) ??
                    Dequeue(TaskPriority.Normal) ??
                    Dequeue(TaskPriority.None);
            }

            return current;
        }

        void IEventLoop.Spin()
        {
        }

        void IEventLoop.CancelAll()
        {
            lock (_lockObj)
            {
                _queues.Clear();
                _cts?.Cancel();
            }
        }

        private void Runner(Object state)
        {
            if (state is CancellationToken token)
            {
                while (!token.IsCancellationRequested)
                {
                    var current = Dequeue();

                    if (current == null)
                    {
                        Thread.Sleep(1);
                        continue;
                    }

                    using (token.Register(current.Cancel))
                    {
                        current.Run();
                    }
                }
            }
        }

        void IDisposable.Dispose() => _cts?.Cancel();

        private class LoopEntry : ICancellable
        {
            private readonly CancellationTokenSource cts = new CancellationTokenSource();
            private readonly Action<CancellationToken> _action;
            private readonly Action<Exception> _trackError;

            public LoopEntry(Action<CancellationToken> action, Action<Exception> trackError)
            {
                _action = action;
                _trackError = trackError;
            }

            public Boolean IsCompleted { get; set; } = false;

            public Boolean IsRunning { get; set; } = false;

            public void Run()
            {
                IsCompleted = false;
                IsRunning = true;

                try { _action.Invoke(cts.Token); }
                catch (Exception ex)
                {
                    _trackError?.Invoke(ex);
                }

                IsRunning = false;
                IsCompleted = true;
            }

            public void Cancel() => cts.Cancel();
        }
    }
}
