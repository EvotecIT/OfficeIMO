"""Opt-in fixture generation with msoffcrypto-tool 6.0.0; never used at runtime.

Run beside the paired plaintext workbook. The password is public test data.
The independent engine's SHA-1 integrity generator lacks AES padding, so this
script supplies the specification's zero padding with its existing AES helper.
Pass --password-vectors to generate the Unicode password/key-size fixtures
instead of the package integrity hash variants.
"""

import hashlib
import hmac
import io
from pathlib import Path
import secrets
import sys

from msoffcrypto.method.container.ecma376_encrypted import ECMA376Encrypted
from msoffcrypto.method.ecma376_agile import (
    ECMA376Agile,
    ECMA376AgileEncryptionInfo,
    _decrypt_aes_cbc,
    _encrypt_aes_cbc_padded,
    _generate_iv,
    _normalize_key,
    blkKey_VerifierHashInput,
    blkKey_encryptedVerifierHashValue,
    blkKey_encryptedKeyValue,
    blkKey_dataIntegrity1,
    blkKey_dataIntegrity2,
)


def write_fixture(root, filename, plaintext, info, key):
    params = info.keyData
    name = params.hashName
    payload = ECMA376Agile.encrypt_payload(io.BytesIO(plaintext), params, key, params.saltValue)
    integrity_key = secrets.token_bytes(params.hashSize)
    iv1 = _generate_iv(params, blkKey_dataIntegrity1, params.saltValue)
    iv2 = _generate_iv(params, blkKey_dataIntegrity2, params.saltValue)
    digest = hmac.digest(integrity_key, payload, name.lower())
    info.encryptedHmacKey = _encrypt_aes_cbc_padded(integrity_key, key, iv1, params.blockSize)
    info.encryptedHmacValue = _encrypt_aes_cbc_padded(digest, key, iv2, params.blockSize)

    # Independent AES/HMAC and payload decryption protect the generated
    # fixture's meaningful bytes, including the 8-byte package size header.
    recovered_key = _decrypt_aes_cbc(info.encryptedHmacKey, key, iv1)
    recovered_digest = _decrypt_aes_cbc(info.encryptedHmacValue, key, iv2)
    assert recovered_key[params.hashSize:] == bytes(len(recovered_key) - params.hashSize)
    assert recovered_digest[params.hashSize:] == bytes(len(recovered_digest) - params.hashSize)
    assert hmac.compare_digest(
        hmac.digest(recovered_key[:params.hashSize], payload, name.lower()),
        recovered_digest[:params.hashSize],
    )
    assert ECMA376Agile.decrypt(key, params.saltValue, name, io.BytesIO(payload)) == plaintext

    descriptor = info.getEncryptionDescriptorHeader() + info.toEncryptionDescriptor().encode("utf-8")
    output = io.BytesIO()
    ECMA376Encrypted(payload, descriptor).write_to(output)
    path = root / filename
    path.write_bytes(output.getvalue())
    print(path.name, path.stat().st_size, hashlib.sha256(output.getvalue()).hexdigest().upper())


def password_vector(password, hash_name, key_bits, spin_count):
    # The independent producer fixes its public writer to SHA-512/AES-256.
    # Reuse its password chain and AES helpers with supported descriptor values.
    # Short digests need the MS-OFFCRYPTO 0x36 padding before AES key use.
    info = ECMA376AgileEncryptionInfo()
    info.spinCount = spin_count
    params = info.encryptedKey
    params.hashName = hash_name
    params.hashSize = hashlib.new(hash_name.lower()).digest_size
    params.keyBits = key_bits
    params.saltValue = secrets.token_bytes(params.saltSize)
    chain = ECMA376Agile._derive_iterated_hash_from_password(
        password, params.saltValue, hash_name, spin_count).digest()

    def derive(block_key):
        key = ECMA376Agile._derive_encryption_key(chain, block_key, hash_name, key_bits)
        return _normalize_key(key, key_bits // 8)

    verifier = secrets.token_bytes(params.saltSize)
    verifier_hash = hashlib.new(hash_name.lower(), verifier).digest()
    info.encryptedVerifierHashInput = _encrypt_aes_cbc_padded(
        verifier, derive(blkKey_VerifierHashInput), params.saltValue, params.blockSize)
    info.encryptedVerifierHashValue = _encrypt_aes_cbc_padded(
        verifier_hash, derive(blkKey_encryptedVerifierHashValue), params.saltValue, params.blockSize)
    secret_key = secrets.token_bytes(info.keyData.keyBits // 8)
    info.encryptedKeyValue = _encrypt_aes_cbc_padded(
        secret_key, derive(blkKey_encryptedKeyValue), params.saltValue, params.blockSize)
    info.keyData.saltValue = secrets.token_bytes(info.keyData.saltSize)

    # Verify the password's full digest and zero padding before publishing.
    assert _decrypt_aes_cbc(info.encryptedVerifierHashInput,
        derive(blkKey_VerifierHashInput), params.saltValue) == verifier
    recovered = _decrypt_aes_cbc(info.encryptedVerifierHashValue,
        derive(blkKey_encryptedVerifierHashValue), params.saltValue)
    assert hmac.compare_digest(recovered[:params.hashSize], verifier_hash)
    assert recovered[params.hashSize:] == bytes(len(recovered) - params.hashSize)
    assert _decrypt_aes_cbc(info.encryptedKeyValue,
        derive(blkKey_encryptedKeyValue), params.saltValue) == secret_key
    return info, secret_key


def main():
    root = Path(__file__).resolve().parent
    plaintext = (root / "agile-aes256-sha512.plain.xlsx").read_bytes()
    if "--password-vectors" in sys.argv[1:]:
        vectors = (
            ("agile-password-sha1-aes256.xlsx", "Zażółć-🙂-密碼", "SHA1", 256, 2049),
            ("agile-password-sha256-aes192.xlsx", "Другой-🔑-päss", "SHA256", 192, 17),
        )
        for filename, password, hash_name, key_bits, spin_count in vectors:
            info, key = password_vector(password, hash_name, key_bits, spin_count)
            write_fixture(root, filename, plaintext, info, key)
    else:
        for name in ("SHA1", "SHA256", "SHA384", "SHA512"):
            info, key = ECMA376Agile.generate_encryption_parameters("hunter2", spin_count=1000)
            info.keyData.hashName = name
            info.keyData.hashSize = hashlib.new(name.lower()).digest_size
            write_fixture(root, "agile-digest-" + name.lower() + ".xlsx", plaintext, info, key)


if __name__ == "__main__":
    main()
