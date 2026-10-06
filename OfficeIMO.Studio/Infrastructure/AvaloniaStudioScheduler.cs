using Avalonia.Threading;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Runs shared session callbacks on the owning Avalonia dispatcher.</summary>
internal sealed class AvaloniaStudioScheduler : IStudioScheduler {
    private readonly Dispatcher _dispatcher = Dispatcher.UIThread;
    public void Post(Action callback) => _dispatcher.Post(callback);
    public IDisposable Schedule(TimeSpan delay, Action callback) => new ScheduledCallback(delay, callback);
    private sealed class ScheduledCallback : IDisposable {
        private readonly DispatcherTimer _timer;
        internal ScheduledCallback(TimeSpan delay, Action callback) {
            _timer = new() { Interval = delay };
            _timer.Tick += (_, _) => { Dispose(); callback(); };
            _timer.Start();
        }
        public void Dispose() => _timer.Stop();
    }
}
