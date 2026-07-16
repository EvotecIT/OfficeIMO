using Google.Apis.Util.Store;

namespace OfficeIMO.GoogleWorkspace.Auth.GoogleApis {
    /// <summary>
    /// Persistence boundary for OAuth refresh tokens and related authorization state.
    /// Implementations are responsible for encrypting sensitive values at rest.
    /// </summary>
    public interface IGoogleWorkspaceTokenStore {
        Task StoreAsync<T>(string key, T value);
        Task DeleteAsync<T>(string key);
        Task<T?> GetAsync<T>(string key);
        Task ClearAsync();
    }

    /// <summary>
    /// Bridges an OfficeIMO token store to the data-store contract used by Google.Apis.Auth.
    /// </summary>
    public sealed class GoogleApisDataStoreAdapter : IDataStore {
        private readonly IGoogleWorkspaceTokenStore _store;

        public GoogleApisDataStoreAdapter(IGoogleWorkspaceTokenStore store) {
            _store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public Task StoreAsync<T>(string key, T value) => _store.StoreAsync(key, value);
        public Task DeleteAsync<T>(string key) => _store.DeleteAsync<T>(key);
        public Task<T> GetAsync<T>(string key) => GetRequiredAsync<T>(key);
        public Task ClearAsync() => _store.ClearAsync();

        private async Task<T> GetRequiredAsync<T>(string key) {
            T? value = await _store.GetAsync<T>(key).ConfigureAwait(false);
            return value!;
        }
    }
}
