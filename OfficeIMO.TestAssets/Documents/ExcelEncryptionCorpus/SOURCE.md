# Office Agile encryption fixtures

These XLSX and XLSB files exercise Agile AES-256/SHA-512 decryption and the
forward-only encrypted workbook readers. The test password is `hunter2`.

The files are copied unchanged from the
[ExcelReader test corpus at ca5b50f99e8ef57ab476f0a2bc8043558d58b28d](https://github.com/GabrielMarquezMatte/ExcelReader/tree/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Tests/data/encrypted).
The original MIT license is included in `LICENSE.txt`.

The source corpus documents the paired plaintext files as outputs of the
independent `msoffcrypto-tool` decryptor. They provide a byte-for-byte decryption
oracle. The encryption producer is not pinned separately for these Agile files,
so they do not establish independent encryption-producer provenance. The tools
named here describe fixture provenance and are not OfficeIMO runtime dependencies.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| agile-aes256-sha512.xlsx | 15360 | 13F9532B916634CB112D3593E7AB859F67B2DA35A3A47E574F1C1E26C5DA370D |
| agile-aes256-sha512.plain.xlsx | 8915 | 85FBDE3D5CC6C936BE8D9C8A5F9658E4CDED25145A2E6449C5D1FCD54F64E602 |
| agile-aes256-sha512.xlsb | 14848 | 295B555E411A3900B43DF0753A1965367C861B9E290B72B912E2FD511BDE1321 |
| agile-aes256-sha512.plain.xlsb | 8221 | 3D0A3715B4E1D0749DB110B9B7E8A6FEE604A158DAE09EE49B0566267E426C7C |

The `agile-digest-*` XLSX files are generated from the paired plaintext workbook
with the independent [msoffcrypto-tool 6.0.0 engine](https://github.com/nolze/msoffcrypto-tool/blob/v6.0.0/msoffcrypto/method/ecma376_agile.py).
The reproducible procedure is
in `generate-integrity-fixtures.py`; running it requires that optional test tool
and produces new random salts and keys. Password derivation uses SHA-512 with
1,000 iterations; package encryption and integrity use the hash named in the file.
The HMAC key has the digest length, with zero padding for SHA-1's 20-byte key and
digest. The script verifies the complete encrypted package HMAC and recovered
plaintext using Python HMAC and the independent engine's AES/decryption routines.
These establish independently generated cryptographic interoperability, not
Microsoft Office producer provenance.

[MS-OFFCRYPTO section 2.3.4.14](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/63d9c262-82b9-4fa3-a06d-d087b93e3b00)
specifies a salt-size HMAC key. Excel 16.0 and the independent engine use a
digest-size HMAC key. OfficeIMO writes the interoperable digest-size representation
and reads both it and existing salt-size output. Reader tests cover both
representations and reject modified ciphertext, tags, unknown key lengths, and
nonzero key padding. The generator does not run in normal builds or tests.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| agile-digest-sha1.xlsx | 14336 | 5B47AB12169D8D5D790B2A2E02FEB1069D0D55E293D2635D5008D29E41C3BD7D |
| agile-digest-sha256.xlsx | 14336 | ECFBB09871C6AAD780192342B443C7515E8BE37FD84AFA77F3D33DED92444A95 |
| agile-digest-sha384.xlsx | 14336 | A58C9F8FB45173356E42BC20A8E6258A93009E376F185D722346F34D46862F96 |
| agile-digest-sha512.xlsx | 14336 | 74C4337D90AD249B126ED48A281968329F12C7352AAB73C736E172C57C2FB94B |

The `officeimo-salt-*` files preserve the previous OfficeIMO HMAC layout for
compatibility tests. They contain the same paired plaintext workbook, encrypted
with the Core engine before its digest-size writer correction, AES-128, one
password iteration, and the hash named in each file. Their public test password
is `test-password`. The directory uses the corrected MS-CFB pointer and ordering
contract from commit `9a6256caeafff96322ab7f71e333371f18d3007c`; the encrypted payload,
password verifier, and salt-size HMAC generation remain the previous engine's
behavior. An independent decryptor verified their entire package HMAC and exact
plaintext. These are compatibility fixtures, not independent-producer evidence.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| officeimo-salt-sha1.xlsx | 12800 | 597D579052DC7B2485E99DF70A7C0437F85DEFE22DA3E872D9D6D338B18F633F |
| officeimo-salt-sha256.xlsx | 12800 | 0B340D2883BADA479564C37E3582802E08D295E91A1A10141D15E55B86721D05 |
| officeimo-salt-sha384.xlsx | 12800 | 728C0CE8F95C2B9E29F542E67CE43F44EC43C06D77E9399E60BA2BB46617DE4B |
| officeimo-salt-sha512.xlsx | 12800 | E93DE625A4D7B93E944F3C07D51AB9FA9384D31BAFFDB3A75D763F2455999C11 |

The `agile-password-*` fixtures exercise password derivation separately from
package encryption. Run `generate-integrity-fixtures.py --password-vectors` to
generate them with the independent engine's password hash and AES helpers. Both
contain the same paired plaintext and use AES-256/SHA-512 for the package and
integrity. Their password descriptors use the following public test data:

| File | Password | Password hash | Password AES key | Iterations |
| --- | --- | --- | ---: | ---: |
| agile-password-sha1-aes256.xlsx | `Zażółć-🙂-密碼` | SHA-1 | 256 bits | 2049 |
| agile-password-sha256-aes192.xlsx | `Другой-🔑-päss` | SHA-256 | 192 bits | 17 |

The first vector covers UTF-16 encoding, repeated counter hashing, and the
[specified `0x36` padding](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/74d60145-a0f0-44be-99ce-c65d211b4eb7)
when a hash is shorter than the AES key. The second
covers truncation to a 192-bit password key. The generator verifies the password
verifier, key recovery, complete HMAC, and exact plaintext before writing either
file. Its SHA-1 key padding extends the independent engine's key helper, which
otherwise only truncates short digests. These are independently generated test
vectors rather than Microsoft Office producer evidence.

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| agile-password-sha1-aes256.xlsx | 14336 | 56EF3FA8F2FA7FC39FA70CB87C7C34845405B939151399930E5EBD547318C777 |
| agile-password-sha256-aes192.xlsx | 14336 | 6650F541BA13727FA07C4202C404EB590DF68EC83997F80B2EBCF415938BCEC0 |
