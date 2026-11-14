"""Writes public-key encrypted PDF fixtures with pyHanko.

Usage: python generate_pyhanko_pubkey.py <output-directory>

Creates a throwaway RSA recipient (recipient.pfx, password "officeimo-test"),
a one-page plain PDF, and one encrypted copy per security handler variant.
"""
import datetime
import io
import sys
from pathlib import Path

from asn1crypto import x509 as asn1_x509
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.x509.oid import NameOID
from pyhanko.pdf_utils.crypt import PubKeySecurityHandler, SecurityHandlerVersion
from pyhanko.pdf_utils.crypt.permissions import PubKeyPermissions
from pyhanko.pdf_utils.reader import PdfFileReader
from pyhanko.pdf_utils.writer import copy_into_new_writer

TEXT = "Public key fixture"


def plain_pdf() -> bytes:
    content = ("BT /F1 18 Tf 72 720 Td (" + TEXT + ") Tj ET").encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
        b"<< /Title (Public key fixture title) >>",
    ]
    out = io.BytesIO()
    out.write(b"%PDF-1.7\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(out.tell())
        out.write(str(number).encode() + b" 0 obj\n" + body + b"\nendobj\n")
    xref = out.tell()
    out.write(b"xref\n0 " + str(len(objects) + 1).encode() + b"\n0000000000 65535 f \n")
    for offset in offsets:
        out.write(b"%010d 00000 n \n" % offset)
    out.write(b"trailer\n<< /Size " + str(len(objects) + 1).encode() + b" /Root 1 0 R /Info 6 0 R"
              b" /ID [<0123456789abcdef0123456789abcdef> <0123456789abcdef0123456789abcdef>] >>\n"
              b"startxref\n" + str(xref).encode() + b"\n%%EOF\n")
    return out.getvalue()


def recipient(output: Path):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "OfficeIMO public-key PDF recipient")])
    now = datetime.datetime(2026, 1, 1, tzinfo=datetime.timezone.utc)
    certificate = (
        x509.CertificateBuilder()
        .subject_name(name).issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now).not_valid_after(now + datetime.timedelta(days=36500))
        .add_extension(x509.KeyUsage(digital_signature=False, content_commitment=False, key_encipherment=True,
                                     data_encipherment=False, key_agreement=False, key_cert_sign=False,
                                     crl_sign=False, encipher_only=False, decipher_only=False), critical=True)
        .sign(key, hashes.SHA256())
    )
    pfx = pkcs12.serialize_key_and_certificates(
        b"recipient", key, certificate, None,
        serialization.BestAvailableEncryption(b"officeimo-test"))
    (output / "recipient.pfx").write_bytes(pfx)
    return asn1_x509.Certificate.load(certificate.public_bytes(serialization.Encoding.DER))


def encrypt(source: bytes, cert, path: Path, **kwargs) -> None:
    writer = copy_into_new_writer(PdfFileReader(io.BytesIO(source)))
    handler = PubKeySecurityHandler.build_from_certs([cert], **kwargs)
    writer._assign_security_handler(handler)
    buffer = io.BytesIO()
    writer.write(buffer)
    path.write_bytes(buffer.getvalue())


def main() -> None:
    output = Path(sys.argv[1])
    output.mkdir(parents=True, exist_ok=True)
    cert = recipient(output)
    source = plain_pdf()
    no_print = PubKeyPermissions.allow_everything() & ~PubKeyPermissions.ALLOW_PRINTING & ~PubKeyPermissions.ALLOW_HIGH_QUALITY_PRINTING
    encrypt(source, cert, output / "pyhanko-s4-rc4-128.pdf", version=SecurityHandlerVersion.RC4_LONGER_KEYS,
            keylen_bytes=16, use_aes=False, use_crypt_filters=False, pdf_mac=False)
    encrypt(source, cert, output / "pyhanko-s5-aesv2.pdf", version=SecurityHandlerVersion.RC4_OR_AES128,
            keylen_bytes=16, use_aes=True, pdf_mac=False)
    encrypt(source, cert, output / "pyhanko-s5-aesv3.pdf", version=SecurityHandlerVersion.AES256,
            keylen_bytes=32, use_aes=True, pdf_mac=False, perms=no_print)
    encrypt(source, cert, output / "pyhanko-s5-aesv3-plain-metadata.pdf", version=SecurityHandlerVersion.AES256,
            keylen_bytes=32, use_aes=True, pdf_mac=False, encrypt_metadata=False)
    # ISO/TS 32003 AES-GCM (/V 6); OfficeIMO.Pdf must refuse it rather than misread it.
    encrypt(source, cert, output / "pyhanko-s5-aes-gcm.pdf", version=SecurityHandlerVersion.AES_GCM,
            keylen_bytes=32, use_aes=True, pdf_mac=False)


if __name__ == "__main__":
    main()
