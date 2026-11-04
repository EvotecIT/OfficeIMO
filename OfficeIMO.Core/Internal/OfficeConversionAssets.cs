using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

/// <summary>Defers conversion sidecars until their owner accepts the completed report.</summary>
internal sealed class OfficeConversionAssets : IDisposable {
    private static readonly AsyncLocal<OfficeConversionAssets?> Current = new();
    private readonly OfficeConversionAssets? _parent;
    private readonly List<Asset> _assets = new();
    private readonly List<string> _directories = new();
    private readonly CancellationToken _token;
    private bool _completed;

    internal OfficeConversionAssets(CancellationToken token) {
        _token = token;
        _parent = Current.Value;
        Current.Value = this;
    }

    internal IReadOnlyList<Asset> Pending => _assets;

    internal static void WriteFile(string path, Action<Stream> write) {
        OfficeConversionAssets? scope = Current.Value;
        if (scope == null) { OfficeFileCommit.Write(path, write); return; }
        scope._token.ThrowIfCancellationRequested();
        if (scope._directories.Count == 0) scope._directories.Add(OfficeTemporaryDirectory.Create("officeimo-conversion-assets-"));
        string staged = Path.Combine(scope._directories[0], Guid.NewGuid().ToString("N"));
        using (var stream = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None)) write(stream);
        scope._assets.Add(new Asset(Path.GetFullPath(path), staged));
    }

    internal void Complete() {
        if (TransferToParent()) return;
        foreach (Asset asset in _assets) {
            _token.ThrowIfCancellationRequested();
            OfficeFileCommit.Write(asset.TargetPath, output => {
                using var input = File.OpenRead(asset.StagedPath);
                input.CopyTo(output);
            }, _token);
        }
        _completed = true;
    }

    internal async Task CompleteAsync(Func<Asset, CancellationToken, Task> publish) {
        if (TransferToParent()) return;
        foreach (Asset asset in _assets) {
            _token.ThrowIfCancellationRequested();
            await publish(asset, _token).ConfigureAwait(false);
        }
        _completed = true;
    }

    private bool TransferToParent() {
        if (_completed) throw new InvalidOperationException("Conversion assets have already been accepted.");
        _token.ThrowIfCancellationRequested();
        if (_parent == null) return false;
        _parent._assets.AddRange(_assets);
        _parent._directories.AddRange(_directories);
        _directories.Clear();
        _completed = true;
        return true;
    }

    public void Dispose() {
        Current.Value = _parent;
        foreach (string directory in _directories) Directory.Delete(directory, recursive: true);
    }

    internal sealed class Asset {
        internal Asset(string targetPath, string stagedPath) { TargetPath = targetPath; StagedPath = stagedPath; }
        internal string TargetPath { get; }
        internal string StagedPath { get; }
    }
}
