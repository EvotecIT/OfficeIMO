using Google.Apis.Util.Store;

namespace OfficeIMO.GoogleWorkspace.Auth.GoogleApis {
    /// <summary>
    /// Persistence boundary for OAuth refresh tokens and related authorization state.
    /// Implementations are responsible for encrypting sensitive values at rest.
    /// </summary>
    public interface IGoogleWorkspaceTokenStore {
        /// <summary>Stores or replaces authorization state under a key.</summary>
        Task StoreAsync<T>(string key, T value);
        /// <summary>Deletes authorization state of the requested type under a key.</summary>
        Task DeleteAsync<T>(string key);
        /// <summary>Reads authorization state, returning <see langword="null"/> when the key is absent.</summary>
        Task<T?> GetAsync<T>(string key);
        /// <summary>Deletes all authorization state owned by this store.</summary>
        Task ClearAsync();
    }

    /// <summary>
    /// Bridges an OfficeIMO token store to the data-store contract used by Google.Apis.Auth.
    /// </summary>
    public sealed class GoogleApisDataStoreAdapter : IDataStore {
        private readonly IGoogleWorkspaceTokenStore _store;

        /// <summary>Creates an adapter over an application-owned secure token store.</summary>
        /// <param name="store">Token store to expose through Google's data-store contract.</param>
        public GoogleApisDataStoreAdapter(IGoogleWorkspaceTokenStore store) {
            _store = store ?? throw new ArgumentNullException(nameof(store));
        }

        /// <inheritdoc />
        public Task StoreAsync<T>(string key, T value) => _store.StoreAsync(key, value);
        /// <inheritdoc />
        public Task DeleteAsync<T>(string key) => _store.DeleteAsync<T>(key);
        /// <inheritdoc />
        public Task<T> GetAsync<T>(string key) => GetRequiredAsync<T>(key);
        /// <inheritdoc />
        public Task ClearAsync() => _store.ClearAsync();

        private async Task<T> GetRequiredAsync<T>(string key) {
            T? value = await _store.GetAsync<T>(key).ConfigureAwait(false);
            return value!;
        }
    }
}
