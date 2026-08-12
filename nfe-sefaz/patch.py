"""Corrige o nfe_downloader.py no local. Execute uma vez."""
with open('nfe_downloader.py', 'r', encoding='utf-8') as f:
    c = f.read()

c = c.replace(
    "from cryptography.x509 import load_pem_x509_certificate\n        from cryptography.hazmat.primitives.serialization import load_pem_private_key",
    "from cryptography.hazmat.primitives.serialization import load_pem_private_key"
)
c = c.replace(
    "\n    cert = load_pem_x509_certificate(cert_pem)\n    key  = load_pem_private_key(key_pem, password=None)",
    "\n    key = load_pem_private_key(key_pem, password=None)"
)
c = c.replace("digest_algorithm='sha1'", "digest_algorithm='sha256'")
c = c.replace("signature_algorithm='rsa-sha1'", "signature_algorithm='rsa-sha256'")
c = c.replace("return signer.sign(xml_element, key=key, cert=cert)", "return signer.sign(xml_element, key=key, cert=cert_pem)")

with open('nfe_downloader.py', 'w', encoding='utf-8') as f:
    f.write(c)

print('Corrigido! Agora rode: py nfe_downloader.py')
