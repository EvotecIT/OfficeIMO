namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Schedules presentation-thread callbacks without depending on a UI framework.</summary>
internal interface IStudioScheduler {
    void Post(Action callback);
    /// <summary>Schedules one callback; disposing the returned handle cancels it.</summary>
    IDisposable Schedule(TimeSpan delay, Action callback);
}
