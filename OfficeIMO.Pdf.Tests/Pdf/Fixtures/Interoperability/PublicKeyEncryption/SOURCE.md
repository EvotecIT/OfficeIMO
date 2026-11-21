# Public-key encrypted PDF fixtures

pyHanko 0.33.0 (with cryptography 50.0.2 and asn1crypto 1.5.1) writes these
PDFs independently of OfficeIMO. Each one is the same one-page document, with the
text `Public key fixture` and the title `Public key fixture title`, encrypted
with the public-key security handler for a single RSA-2048 recipient.

The generator also writes `recipient.pfx`, the recipient's throwaway certificate
and private key (password `officeimo-test`). It protects nothing else. The
repository does not track `.pfx` files, so `PdfPublicKeyEncryptedReadTests`
carries it as a base64 constant (`PyHankoRecipientPkcs12`).

| File | Handler | Notes |
| --- | --- | --- |
| pyhanko-s4-rc4-128.pdf | `adbe.pkcs7.s4`, `/V 2`, RC4 128-bit | Document-level `/Recipients` |
| pyhanko-s5-aesv2.pdf | `adbe.pkcs7.s5`, `/V 4`, AESV2 | Crypt filter recipients |
| pyhanko-s5-aesv3.pdf | `adbe.pkcs7.s5`, `/V 5`, AESV3 | Printing permissions removed for the recipient |
| pyhanko-s5-aesv3-plain-metadata.pdf | `adbe.pkcs7.s5`, `/V 5`, AESV3 | `/EncryptMetadata false` |
| pyhanko-s5-aes-gcm.pdf | `adbe.pkcs7.s5`, `/V 6`, AESV4 (ISO/TS 32003 AES-GCM) | Must be refused as unsupported |

pyHanko is an optional fixture-production tool; ordinary tests consume the
checked-in files and do not require it.

```sh
python3 -m venv venv
./venv/bin/pip install pyhanko==0.33.0
./venv/bin/python generate_pyhanko_pubkey.py /path/to/task-output
```

Each run creates a new key, seed and IVs, so regenerated files differ byte for
byte. When refreshing them, replace all PDFs and update the base64 recipient
constant from the new `recipient.pfx`.

| File | SHA-256 |
| --- | --- |
| pyhanko-s4-rc4-128.pdf | `a5aaaa8899f88847d506e3f8f0c0cbd378aee430909918b7fed07b03efdec38d` |
| pyhanko-s5-aes-gcm.pdf | `960963e6a6e39526d382b21160fc9dccaf1384f0fb6f6a1e1533e4b08a0ee91e` |
| pyhanko-s5-aesv2.pdf | `aa6464b1ab51f75068f9af4216c84804d1fd6497e074cda75f0ad3c33c6dc802` |
| pyhanko-s5-aesv3-plain-metadata.pdf | `5dae7f76fb76b6089befed4c64c9226ce8710f6d101372d52a954b680f79e9e5` |
| pyhanko-s5-aesv3.pdf | `05d89af3220296549d0f3f3bd26cda94b5c546c88b81c01967473c5515045e4d` |
