using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Excel;

public sealed partial class ExcelWorkbookDataReader {
    private DbDataReader? _activeReader;
    // Metadata remains available without opening a worksheet or decoding shared strings.
    private DbDataReader CurrentReader {
        get => GetCurrentReader(_cancellationToken);
        set => _activeReader = value;
    }

    private DbDataReader GetCurrentReader(CancellationToken callerToken) {
        ThrowIfClosed();
        _cancellationToken.ThrowIfCancellationRequested();
        callerToken.ThrowIfCancellationRequested();
        if (_activeReader != null) return _activeReader;
        CancellationTokenSource? sheetSource = null;
        CancellationTokenRegistration registration = default;
        try {
            CancellationToken token = _cancellationToken;
            if (callerToken.CanBeCanceled && callerToken != _cancellationToken) {
                sheetSource = _cancellationToken.CanBeCanceled
                    ? CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken)
                    : new CancellationTokenSource();
                registration = callerToken.Register(static state => ((CancellationTokenSource)state!).Cancel(), sheetSource);
                token = sheetSource.Token;
            }
            _activeReader = _openSheet(_resultIndex, token);
            callerToken.ThrowIfCancellationRequested();
            _cancellationToken.ThrowIfCancellationRequested();
            registration.Dispose();
            _currentSheetCancellationSource = sheetSource;
            sheetSource = null;
            return _activeReader;
        } catch {
            CloseAfterSheetOpenFailure();
            throw;
        } finally {
            registration.Dispose();
            sheetSource?.Dispose();
        }
    }
}
