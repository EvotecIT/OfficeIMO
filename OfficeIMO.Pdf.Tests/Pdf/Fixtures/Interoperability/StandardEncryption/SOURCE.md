# Standard security fixtures

`New-StandardEncryptionFixtures.ps1` writes every PDF in this folder with PowerShell 7.4+ and
`System.Security.Cryptography` alone; no OfficeIMO code takes part. Each file has one page whose
content stream and `/Info /Title` carry the text the tests read. Rerunning the script draws new
keys, salts and nonces, so the bytes change but the contract does not.

| File | Security | User / owner password | What it exercises |
| --- | --- | --- | --- |
| `r6-saslprep.pdf` | `/V 5 /R 6`, AESV3 | `SaSLprep` (stored after SASLprep) / `owner` | RFC 4013 mapping of `S` U+00AA `SL` U+00AD `prep` before UTF-8 encoding |
| `r6-unprepared-password.pdf` | `/V 5 /R 6`, AESV3 | `pass` U+00AD `word` (stored without SASLprep) / `owner` | Producers that skipped the SASLprep mapping |
| `r6-owner-only-padded-entries.pdf` | `/V 5 /R 6`, AESV3, `/P -1052` | empty / `owner` | `/O` and `/U` padded to 127 bytes as Acrobat X and XI write them; only the first 48 bytes are defined |
| `r6-identity-crypt-filter-stream.pdf` | `/V 5 /R 6`, AESV3 | `open` / `owner` | Page content in clear text behind `/Filter [/Crypt]` with `/Name /Identity` |
| `r6-crypt-filter-without-method.pdf` | `/V 5 /R 6`, AESV3 | `open` / `owner` | Page content in clear text behind a named crypt filter that omits `/CFM` |
| `r6-embedded-files-only.pdf` | `/V 5 /R 6`, `/StmF` and `/StrF /Identity`, `/EFF /StdCF` AESV3 `/AuthEvent /EFOpen` | `open` / `owner` | Readable without a password; only the embedded file is encrypted |
| `r7-aes-gcm.pdf` | `/V 6 /R 7`, AESV4 (AES-256-GCM, ISO/TS 32003) | `open` / `owner` | 12-byte nonce, ciphertext and 16-byte tag per string and stream |
| `r3-rc4-40-default-length.pdf` | `/V 1 /R 3`, RC4 without `/Length` | `open` / `owner` | The 40-bit default key length |
| `r3-rc4-128-without-length.pdf` | `/V 2 /R 3`, 128-bit RC4 without `/Length` | `open` / `owner` | A producer that omitted `/Length` for a 128-bit key |
| `r3-literal-id-with-bracket.pdf` | `/V 1 /R 3`, RC4 | `open` / `owner` | A literal-string `/ID` containing `]` and the raw byte 0xE9 |

The revision 6 and 7 password entries follow ISO 32000-2 algorithm 2.B and algorithms 8 to 10; the
revision 3 entries follow algorithms 2, 3 (whole-digest rounds) and 5. The revision 7 file has no ISO/TS 32004 `/KDFSalt`
or document MAC; OfficeIMO does not verify those yet.
