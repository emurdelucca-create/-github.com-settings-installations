"""Corrige o nfe_downloader.py no local. Execute uma vez."""
with open('nfe_downloader.py', 'r', encoding='utf-8') as f:
    c = f.read()

# Fix 1: remover assinatura XML do SOAP (SEFAZ rejeita cStat 215)
old_soap = '''    dist_xml = etree.fromstring(
        f\'<distDFeInt xmlns="{NS_NFE}" versao="1.01">\'
        f\'<tpAmb>{ambiente}</tpAmb>\'
        f\'<cUFAutor>{cuf}</cUFAutor>\'
        f\'<CNPJ>{cnpj}</CNPJ>\'
        f\'<distNSU><ultNSU>{ult_nsu}</ultNSU></distNSU>\'
        f\'</distDFeInt>\'.encode()
    )
    signed     = assinar_xml(dist_xml, cert_pem, key_pem)
    signed_str = etree.tostring(signed, encoding=\'unicode\')

    return (
        \'<?xml version="1.0" encoding="UTF-8"?>\'
        f\'<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\'
        f\' xmlns:xsd="http://www.w3.org/2001/XMLSchema"\'
        f\' xmlns:soap12="{SOAP_NS}">\'
        f\'<soap12:Header>\'
        f\'<nfeCabecMsg xmlns="{NS_DIST}">\'
        f\'<cUF>{cuf}</cUF><versaoDados>1.01</versaoDados>\'
        f\'</nfeCabecMsg>\'
        f\'</soap12:Header>\'
        f\'<soap12:Body>\'
        f\'<nfeDistDFeInteresse xmlns="{NS_DIST}">\'
        f\'<nfeDadosMsg>{signed_str}</nfeDadosMsg>\'
        f\'</nfeDistDFeInteresse>\'
        f\'</soap12:Body>\'
        f\'</soap12:Envelope>\'
    )'''

new_soap = '''    xml_body = (
        f\'<distDFeInt xmlns="{NS_NFE}" versao="1.01">\'
        f\'<tpAmb>{ambiente}</tpAmb>\'
        f\'<cUFAutor>{cuf}</cUFAutor>\'
        f\'<CNPJ>{cnpj}</CNPJ>\'
        f\'<distNSU><ultNSU>{ult_nsu}</ultNSU></distNSU>\'
        f\'</distDFeInt>\'
    )
    return (
        \'<?xml version="1.0" encoding="UTF-8"?>\'
        f\'<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\'
        f\' xmlns:xsd="http://www.w3.org/2001/XMLSchema"\'
        f\' xmlns:soap12="{SOAP_NS}">\'
        f\'<soap12:Header>\'
        f\'<nfeCabecMsg xmlns="{NS_DIST}">\'
        f\'<cUF>{cuf}</cUF><versaoDados>1.01</versaoDados>\'
        f\'</nfeCabecMsg>\'
        f\'</soap12:Header>\'
        f\'<soap12:Body>\'
        f\'<nfeDistDFeInteresse xmlns="{NS_DIST}">\'
        f\'<nfeDadosMsg>{xml_body}</nfeDadosMsg>\'
        f\'</nfeDistDFeInteresse>\'
        f\'</soap12:Body>\'
        f\'</soap12:Envelope>\'
    )'''

if old_soap in c:
    c = c.replace(old_soap, new_soap)
    print('Fix 1 aplicado: removida assinatura XML do SOAP')
else:
    print('Fix 1 já aplicado ou versão diferente')

# Fix 2: sha1 -> sha256
c = c.replace("digest_algorithm='sha1'", "digest_algorithm='sha256'")
c = c.replace("signature_algorithm='rsa-sha1'", "signature_algorithm='rsa-sha256'")

# Fix 3: lxml truth-test
c = c.replace(
    "cstat  = (_find(ret, 'cStat') or type('', (), {'text': '0'})()).text",
    "e_cstat = _find(ret, 'cStat'); cstat = e_cstat.text if e_cstat is not None else '0'"
)
c = c.replace(
    "xmotiv = (_find(ret, 'xMotivo') or type('', (), {'text': ''})()).text",
    "e_xmot = _find(ret, 'xMotivo'); xmotiv = e_xmot.text if e_xmot is not None else ''"
)

# Fix 4: novas colunas (Origem, BC ICMS ST, Vl. ICMS ST, Endereço Emitente, Endereço Destinatário)
old_cab = """CABECALHO = [
    'Chave NF-e', 'Data Emissão', 'Número NF', 'Série', 'Natureza Op.',
    'CNPJ Emitente', 'Razão Social Emitente', 'UF Emitente',
    'CNPJ Destinatário', 'Razão Social Destinatário', 'UF Destinatário',
    'Nº Item', 'Código', 'EAN', 'Descrição', 'NCM', 'CFOP',
    'Unidade', 'Quantidade', 'Valor Unit.', 'Valor Item', 'Desconto Item',
    'Vl. Produtos NF', 'Vl. Frete NF', 'Vl. Desconto NF', 'Vl. Total NF',
    'ICMS', 'PIS', 'COFINS',
]"""

new_cab = """CABECALHO = [
    'Chave NF-e', 'Data Emissão', 'Número NF', 'Série', 'Natureza Op.',
    'CNPJ Emitente', 'Razão Social Emitente', 'UF Emitente',
    'CNPJ Destinatário', 'Razão Social Destinatário', 'UF Destinatário',
    'Nº Item', 'Código', 'EAN', 'Descrição', 'NCM', 'CFOP',
    'Unidade', 'Quantidade', 'Valor Unit.', 'Valor Item', 'Desconto Item',
    'Origem',
    'Vl. Produtos NF', 'Vl. Frete NF', 'Vl. Desconto NF', 'Vl. Total NF',
    'ICMS', 'PIS', 'COFINS',
    'BC ICMS ST', 'Vl. ICMS ST',
    'Endereço Emitente', 'Endereço Destinatário',
]"""

if old_cab in c:
    c = c.replace(old_cab, new_cab)
    print('Fix 4 aplicado: novas colunas adicionadas ao CABECALHO')
else:
    print('Fix 4: CABECALHO já atualizado ou versão diferente')

with open('nfe_downloader.py', 'w', encoding='utf-8') as f:
    f.write(c)

print('Concluido! Rode: py nfe_downloader.py')
