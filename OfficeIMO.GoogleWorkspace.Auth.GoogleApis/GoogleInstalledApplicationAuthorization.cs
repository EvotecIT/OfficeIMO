using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;

namespace OfficeIMO.GoogleWorkspace.Auth.GoogleApis {
    /// <summary>
    /// Settings for interactive authorization of a desktop or other installed application.
    /// </summary>
    public sealed class GoogleInstalledApplicationAuthorizationOptions {
        public ClientSecrets? ClientSecrets { get; set; }
        public IReadOnlyList<string> Scopes { get; set; } = Array.Empty<string>();
        public string? UserId { get; set; }
        public string? Account { get; set; }
        /// <summary>
        /// Required by <see cref="GoogleInstalledApplicationAuthorization.AuthorizeAsync"/>; resolves the
        /// authorized token's provider-issued identity and grants. Read-only callers can use
        /// <see cref="GoogleInstalledApplicationAuthorization.AuthorizeCredentialAsync"/> without a resolver.
        /// </summary>
        public GoogleWorkspaceCredentialBindingResolver? CredentialBindingResolver { get; set; }
        public IGoogleWorkspaceTokenStore? TokenStore { get; set; }
        public ICodeReceiver? CodeReceiver { get; set; }

        internal void Validate() {
            if (ClientSecrets == null || string.IsNullOrWhiteSpace(ClientSecrets.ClientId)) {
                throw new InvalidOperationException("Installed application client secrets with a client ID are required.");
            }

            if (Scopes == null || Scopes.Count == 0 || Scopes.Any(string.IsNullOrWhiteSpace)) {
                throw new InvalidOperationException("At least one non-empty Google API scope is required.");
            }

            if (string.IsNullOrWhiteSpace(UserId)) {
                throw new InvalidOperationException("A stable local user ID is required for token persistence.");
            }

            if (TokenStore == null) {
                throw new InvalidOperationException(
                    "A token store is required. OfficeIMO does not default OAuth refresh tokens to plaintext files.");
            }
        }
    }

    /// <summary>
    /// Runs Google's installed-application authorization flow with PKCE always enabled.
    /// </summary>
    public static class GoogleInstalledApplicationAuthorization {
        public static async Task<UserCredential> AuthorizeCredentialAsync(
            GoogleInstalledApplicationAuthorizationOptions options,
            CancellationToken cancellationToken = default) {
            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            options.Validate();
            var initializer = new GoogleAuthorizationCodeFlow.Initializer {
                ClientSecrets = options.ClientSecrets,
            };

            return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    initializer,
                    options.Scopes,
                    options.UserId!,
                    usePkce: true,
                    taskCancellationToken: cancellationToken,
                    dataStore: new GoogleApisDataStoreAdapter(options.TokenStore!),
                    codeReceiver: options.CodeReceiver)
                .ConfigureAwait(false);
        }

        /// <summary>
        /// Authorizes and returns a credential source with provider-verified identity evidence.
        /// This convenience method always requires <see cref="GoogleInstalledApplicationAuthorizationOptions.CredentialBindingResolver"/>;
        /// use <see cref="AuthorizeCredentialAsync"/> for a resolver-free read-only credential.
        /// </summary>
        public static async Task<GoogleApisCredentialSource> AuthorizeAsync(
            GoogleInstalledApplicationAuthorizationOptions options,
            CancellationToken cancellationToken = default) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            options.Validate();
            if (options.CredentialBindingResolver == null) {
                throw new InvalidOperationException(
                    "A provider-backed credential binding resolver is required before the authorized credential can be used for guarded mutations.");
            }
            UserCredential credential = await AuthorizeCredentialAsync(options, cancellationToken).ConfigureAwait(false);
            return new GoogleApisCredentialSource(credential, null, options.Account, options.CredentialBindingResolver);
        }
    }
}
