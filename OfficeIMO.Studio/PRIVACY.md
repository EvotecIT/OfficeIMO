# OfficeIMO Studio privacy

OfficeIMO Studio opens and processes documents on your device. Opening, reading,
editing, converting, and saving a document does not send it to Evotec.

## Local application data

Studio stores settings, recent-document references, session information,
reading positions, saved signatures, and enabled recovery snapshots in its
user-scoped application data directory. Recovery snapshots can contain document
contents. Recent-document and session references can contain file paths or
storage-provider bookmarks. Settings provides controls for document history
and recovery data.

Diagnostics remain local. They contain event codes, runtime information,
exception types, and sanitized stack frames. They exclude document contents,
document names and paths, and exception messages.

## Optional document assistant

The document assistant uses the provider and model you select. Remote processing
requires enabling the remote-processing control. When you ask a question using
a remote provider, Studio sends your question and selected document evidence to
that provider. Follow-up requests can include conversation context. Provider
privacy policies, account terms, retention policies, and usage limits apply.

Available connections include ChatGPT, an OpenAI-compatible endpoint, GitHub
Copilot, and a local-model endpoint. Local-model connections require a loopback
HTTP(S) endpoint. Check the local server's behavior before sending sensitive
content; a local gateway can forward requests to another service.

Sign-in and model discovery communicate with the selected service without
sending document evidence. Studio retains supported account credentials in its
application-specific authentication files or session memory. API credentials
entered for compatible endpoints remain in session memory; non-secret provider,
model, and endpoint choices can be retained in settings. Protect the device and
its application-data backups. Use Sign out to end the selected Studio connection.

## External tools and support

The direct desktop edition can use an optional Tesseract installation for OCR
and operating-system command-line tools for printer queue delivery. These tools
process the inputs supplied to the operation and have their own licenses and
privacy behavior. The Mac App Store edition does not discover or download
external OCR programs or start printer command-line tools. Document conversions
use OfficeIMO engines in the application.

Files, screenshots, or logs attached to a public GitHub issue are public. Remove
personal information and credentials before sharing an example. For private
support or privacy questions, contact [support@evotec.pl](mailto:support@evotec.pl).

OfficeIMO’s MIT license and public-source policy cover the software and public
project material. They do not make your documents, credentials, or local
application data public.
