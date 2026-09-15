using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;

namespace OfficeIMO.GoogleWorkspace.Auth.GoogleApis {
    /// <summary>
    /// Settings for interactive authorization of a desktop or other installed application.
    /// </summary>
    public sealed class GoogleInstalledApplicationAuthorizationOptions {
        /// <summary>Gets or sets the installed-application OAuth client identifier and secret.</summary>
        public ClientSecrets? ClientSecrets { get; set; }
        /// <summary>Gets or sets the Google API scopes requested during authorization.</summary>
        public IReadOnlyList<string> Scopes { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets the stable local user key used for token persistence.</summary>
        public string? UserId { get; set; }
        /// <summary>Gets or sets an informational account label for resolver-free read-only credentials.</summary>
        public string? Account { get; set; }
        /// <summary>
        /// Required by <see cref="GoogleInstalledApplicationAuthorization.AuthorizeAsync"/>; resolves the
        /// authorized token's provider-issued identity and grants. Read-only callers can use
        /// <see cref="GoogleInstalledApplicationAuthorization.AuthorizeCredentialAsync"/> without a resolver.
        /// </summary>
        public GoogleWorkspaceCredentialBindingResolver? CredentialBindingResolver { get; set; }
        /// <summary>Gets or sets the required application-owned secure OAuth token store.</summary>
        public IGoogleWorkspaceTokenStore? TokenStore { get; set; }
        /// <summary>Gets or sets the receiver used to obtain the authorization code from the user agent.</summary>
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
        /// <summary>Runs the PKCE installed-application flow and returns Google's native user credential.</summary>
        /// <param name="options">Validated OAuth client, scope, user, token-store, and receiver settings.</param>
        /// <param name="cancellationToken">Token used to cancel interactive authorization.</param>
        /// <returns>The authorized Google user credential.</returns>
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
