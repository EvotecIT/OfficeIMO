# OfficeIMO.Invoicing.KSeF

KSeF v2 token authentication, encrypted online FA(3) submission, asynchronous
status and UPO retrieval for .NET 8 and .NET 10. The package depends on
`OfficeIMO.Invoicing.Validation` and the invoice core, with no additional runtime
packages. Native FA(3) authoring belongs to
[OfficeIMO.Invoicing](../OfficeIMO.Invoicing/README.md#create-polish-fa3).

## Reference the library

For a repository consumer, reference the project:

```xml
<ProjectReference Include="../OfficeIMO.Invoicing.KSeF/OfficeIMO.Invoicing.KSeF.csproj" />
```

The library uses an explicit official environment. `KsefEnvironment.Test` is the
default; `Demo` and `Production` require an explicit selection. An optional
`HttpMessageHandler` and `TimeProvider` support host configuration and offline
integration tests. The client owns and disposes its transport. An injected
handler is trusted: it must honor cancellation and must not redirect requests or
record authentication headers and bodies.

## Submit an online invoice

The example accepts exact FA(3) bytes, a caller-owned credential and a persistence
callback. Load the four pinned FA(3) XSD files with `Fa3SchemaBundle.LoadDirectory`
as described in the [validator guide](../OfficeIMO.Invoicing.Validation/README.md).
The callback stores the context, session reference, invoice reference and exact
plaintext hash before closing or waiting for acceptance. It must not store tokens
or session keys.

```csharp
using OfficeIMO.Invoicing.KSeF;
using OfficeIMO.Invoicing.Validation;

public static class OnlineInvoiceExample {
    public static async Task<KsefReceipt> SubmitAsync(
        KsefClient client, KsefContext context, KsefSecret token,
        byte[] fa3Xml, Fa3SchemaValidator validator,
        Action<KsefSubmission> saveReference, CancellationToken cancellationToken) {
        using KsefAuthentication authentication =
            await client.BeginTokenAuthenticationAsync(context, token, cancellationToken);
        while (true) {
            KsefStatus state = await client.GetAuthenticationStatusAsync(authentication, cancellationToken);
            if (state.IsSuccessful) break;
            if (!state.IsPending) throw new InvalidOperationException("KSeF authentication failed.");
            await Task.Delay(TimeSpan.FromSeconds(2), cancellationToken);
        }
        using KsefCredentials credentials = await client.RedeemAsync(authentication, cancellationToken);
        using KsefOnlineSession session = await client.OpenOnlineSessionAsync(credentials, cancellationToken);
        KsefSubmission submission = await client.SubmitAsync(credentials, session, fa3Xml, validator, cancellationToken);
        saveReference(submission);
        await client.CloseOnlineSessionAsync(credentials, session, cancellationToken);
        while (true) {
            KsefInvoiceStatus state = await client.GetInvoiceStatusAsync(credentials, submission, cancellationToken);
            if (state.IsAccepted) break;
            if (!state.Status.IsPending) throw new InvalidOperationException("KSeF invoice processing failed.");
            await Task.Delay(TimeSpan.FromSeconds(2), cancellationToken);
        }
        return await client.GetInvoiceReceiptAsync(credentials, submission, cancellationToken);
    }
}
```

Supply a finite cancellation deadline around the workflow. Persist the known
session reference before submission when recovery must cover a lost submission
response. `ResumeSubmission` recreates a status/receipt handle after a restart
with newly authenticated credentials; it does not reopen a session or resubmit
the invoice. `GetSessionInvoicesAsync` returns one bounded page for matching an
ambiguous submission by its exact plaintext hash. Handle continuation tokens
explicitly and require a unique match before choosing an invoice reference.

## Authentication, encryption and mutation contracts

- Token authentication uses the server challenge's Unix-millisecond timestamp
  and RSA OAEP SHA-256. Certificate and SPKI SHA-256 selectors, certificate
  validity and intended key usage are verified before encryption. This path
  requires a usable RSA certificate; it does not silently substitute EC
  encryption or XAdES certificate authentication.
- Access and refresh tokens have expiration checks. `RefreshAsync` serializes
  refresh operations. Handles are bound to their originating client and context,
  preventing accidental mixing of environments or identities.
- Each online session owns a fresh 32-byte AES key and 16-byte IV. Invoice bytes
  use AES-256 CBC with PKCS#7 padding. Both plaintext and ciphertext sizes and
  SHA-256 digests are sent with the payload. The FA(3) XSD gate runs over the
  captured plaintext before any submission request.
- Redemption is attempted once per authentication handle. Submission and close
  are serialized per session. Connection loss, cancellation after dispatch,
  HTTP 408/5xx, oversized successful responses and malformed successful responses
  produce `KsefMutationAmbiguousException`. No mutation is automatically retried.
  An ambiguous submission or close stops subsequent mutation of that session
  handle; read-only status and page retrieval remain available.
- Other HTTP failures produce `KsefApiException` with status and optional
  `RetryAfter`, without echoing remote error bodies. Read-only calls also have no
  automatic retry policy. The application owns polling and rate-limit policy.
- `KsefSecret`, credentials and sessions clear their owned secret/key buffers on
  disposal. The caller's original strings and runtime HTTP header strings cannot
  be erased by the library. Disposal does not cancel a request already dispatched
  or close a remote session.

The local bounds are 16 MiB for invoice plaintext, 1 MiB for JSON responses,
2 MiB for UPO, 64 advertised certificates and 1,000 invoices per requested page.
Response reading and JSON/XML depth are bounded. The default whole-request
deadline is 30 seconds, configurable from one second to five minutes. These are
library limits; authenticated KSeF limits and permissions can be stricter.

## Acceptance and receipt evidence

| Stage | Evidence |
| --- | --- |
| Native FA(3) XSD | Exact invoice bytes conform to the pinned schema. |
| Submission | The API returned an invoice operation reference. |
| Invoice status | Code 200 and an actual KSeF number, tied to the submitted reference and plaintext hash. |
| UPO retrieval | Direct authenticated official API route; remote presigned URLs are not followed. |
| UPO binding | Exact context, session, KSeF number and invoice hash match uniquely. |
| UPO schema | Separate validation against the unmodified pinned v4-3 XSD. |

`KsefReceipt` retains exact receipt bytes and exposes `RetrievedThroughAuthenticatedApi`,
`IsBound`, `SchemaStatus` and diagnostics separately. Offline `Inspect` never
claims authenticated retrieval. The library does not verify a UPO digital
signature or certify fiscal treatment.

The official pinned TEST invoice/session examples use receiver name
`Ministerstwo Finansów - środowisko testowe (TE)`, while the pinned v4-3 schema
fixes that field to `Ministerstwo Finansów`. Those unmodified examples have valid
identity/hash bindings but an invalid schema result. Neither the fixture nor the
schema is rewritten to conceal the disagreement.

## Qualification harness

Default tests use synthetic HTTP responses and generated RSA keys. They cover
token encryption, AES payloads, exact-byte snapshot capture, asynchronous status,
one-time redemption, cancellation and unknown outcomes, concurrency, bounds,
context ownership, receipt binding and pinned independent receipt bytes.

```powershell
dotnet test OfficeIMO.Invoicing.KSeF.Tests/OfficeIMO.Invoicing.KSeF.Tests.csproj
```

For a read-only official TEST certificate check, set
`OFFICEIMO_KSEF_PUBLIC_TEST=1` and select `PublicTestEncryptionCertificateContract`.
This check needs no identity and makes no authentication or submission request.

Authenticated TEST qualification is a separate opt-in test. Set all of:

| Environment variable | Value |
| --- | --- |
| `OFFICEIMO_KSEF_LIVE_TEST` | `1` to enable the authenticated TEST test. |
| `OFFICEIMO_KSEF_TEST_TOKEN_FILE` | Local UTF-8 file containing the TEST KSeF token; no token goes on the command line. |
| `OFFICEIMO_KSEF_TEST_NIP` | Authorized TEST context NIP. |
| `OFFICEIMO_KSEF_TEST_INVOICE_FILE` | Schema-valid, unique TEST FA(3) invoice XML. |
| `OFFICEIMO_KSEF_TEST_SCHEMA_DIRECTORY` | Directory with the four pinned FA(3) XSD files. |
| `OFFICEIMO_KSEF_TEST_RECEIPT_DIRECTORY` | Caller-selected directory for nonsecret recovery references and exact UPO evidence. |

Run with `--filter FullyQualifiedName~AuthenticatedTestSubmissionStatusAndReceiptQualification`.
The harness uses TEST only, bounds polling and input files, records references
before waiting, creates evidence files without overwriting existing files, and
never replays an ambiguous mutation. A missing identity leaves this test skipped;
offline tests and the public certificate check do not prove authenticated
acceptance or fiscal correctness.

## Protocol sources

The contract uses the Ministry of Finance
[KSeF integrator documentation](https://github.com/CIRFMF/ksef-api/tree/c50f855ef3ae550396c6a272594047506e220498),
its OpenAPI 2.8.1 TE snapshot, authentication schema 2.1 and UPO schema v4-3.
Embedded schema bytes are verified by SHA-256. Exact identities, source paths and
the upstream MIT notice are retained in `ThirdPartyNotices.txt` and `Schemas/LICENSE.txt`.
XAdES/certificate authentication, batch/offline submission, token/permission
administration and receipt signature verification require separate integration
contracts.
