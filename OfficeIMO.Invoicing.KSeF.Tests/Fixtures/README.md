# Independent KSeF receipt fixtures

These are unmodified Ministry of Finance examples from KSeF documentation commit
`c50f855ef3ae550396c6a272594047506e220498`. Filenames are shortened; bytes remain
identical to the upstream files. They use fictional TEST data.

| Local file | Upstream path under `faktury/upo/przyklady/v4-3/kontekst-nip/` | SHA-256 |
| --- | --- | --- |
| `upo-invoice-nip.xml` | `upo-faktura-kontekst-id-nip.xml` | `DA5688B6DB93181E945ACAEDFAF55F8BBBA9197B5F35751327501B1A87BD4422` |
| `upo-session-nip.xml` | `upo-sesja-kontekst-id-nip.xml` | `AB42D4090C3E453935E48FFB450F6FD10911110DA99EB9500A75E296A4C1FDBB` |

Source: [official pinned examples](https://github.com/CIRFMF/ksef-api/tree/c50f855ef3ae550396c6a272594047506e220498/faktury/upo/przyklady/v4-3/kontekst-nip).
The upstream MIT license is retained at
`../../OfficeIMO.Invoicing.KSeF/Schemas/LICENSE.txt` from the test project root.

The examples' TEST receiver name differs from the schema's fixed receiver name.
Tests retain this invalid XSD result while proving independent context/session/
document-hash binding. Synthetic harness responses are derived separately and
are not presented as independent server-produced receipt evidence.
