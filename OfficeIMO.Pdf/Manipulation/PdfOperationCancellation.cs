namespace OfficeIMO.Pdf;

/// <summary>Options whose operation reads cancellation from an internal per-call copy.</summary>
internal interface IPdfOperationCancellation {
    System.Threading.CancellationToken CancellationToken { get; set; }
    object CloneForOperation();
}

/// <summary>Binds a method's cancellation token to the options its operation reads.</summary>
internal static class PdfOperationCancellation {
    /// <summary>
    /// Returns a per-call copy carrying <paramref name="cancellationToken"/>, so the caller's options are not changed.
    /// A token that cannot be canceled keeps the options as they are; optional null options get defaults.
    /// </summary>
    internal static T? Bind<T>(T? options, bool createWhenNull, System.Threading.CancellationToken cancellationToken)
        where T : class, IPdfOperationCancellation, new() {
        if (!cancellationToken.CanBeCanceled) return options;
        if (options == null && !createWhenNull) return null;
        T copy = options == null ? new T() : (T)options.CloneForOperation();
        copy.CancellationToken = cancellationToken;
        return copy;
    }

    /// <summary>Binds the token to a per-call copy of options that have no default instance; null stays null.</summary>
    internal static T? Bind<T>(T? options, System.Threading.CancellationToken cancellationToken)
        where T : class, IPdfOperationCancellation {
        if (options == null || !cancellationToken.CanBeCanceled) return options;
        var copy = (T)options.CloneForOperation();
        copy.CancellationToken = cancellationToken;
        return copy;
    }
}
