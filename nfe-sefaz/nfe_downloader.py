#!/usr/bin/env python3
"""
NF-e Downloader — SEFAZ DistribuicaoDFe
========================================
Baixa todas as NF-e do CNPJ configurado via API da SEFAZ Nacional
e grava os itens detalhados em uma planilha do Google Sheets.

Abas geradas:
  - Compras : NF-e emitidas por terceiros para o seu CNPJ
  - Vendas  : NF-e emitidas pelo seu CNPJ

Na segunda execução em diante, só baixa os documentos novos
(a partir do último NSU salvo em config.json).
"""

import base64
import gzip
import json
import os
import re
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

import requests
from lxml import etree
from cryptography.hazmat.primitives.serialization.pkcs12 import load_key_and_certificates
from cryptography.hazmat.primitives.serialization import (
    Encoding, PrivateFormat, NoEncryption,
)
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────────────────────
CONFIG_FILE = 'config.json'
SEFAZ_URL   = 'https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx'
NS_DIST     = 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe'
NS_NFE      = 'http://www.portalfiscal.inf.br/nfe'
SOAP_NS     = 'http://www.w3.org/2003/05/soap-envelope'
SHEETS_SCOPE = ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive']

CABECALHO = [
    'Chave NF-e', 'Data Emissão', 'Número NF', 'Série', 'Natureza Op.',
    'CNPJ Emitente', 'Razão Social Emitente', 'UF Emitente',
    'CNPJ Destinatário', 'Razão Social Destinatário', 'UF Destinatário',
    'Nº Item', 'Código', 'EAN', 'Descrição', 'NCM', 'CFOP',
    'Unidade', 'Quantidade', 'Valor Unit.', 'Valor Item', 'Desconto Item',
    'Vl. Produtos NF', 'Vl. Frete NF', 'Vl. Desconto NF', 'Vl. Total NF',
    'ICMS', 'PIS', 'COFINS',
]


# ─────────────────────────────────────────────────────────────
# Utilitários
# ─────────────────────────────────────────────────────────────
def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


def carregar_config():
    if not Path(CONFIG_FILE).exists():
        print(f"\n❌ Arquivo '{CONFIG_FILE}' não encontrado.")
        print("   Copie 'config.exemplo.json' para 'config.json' e preencha os dados.\n")
        sys.exit(1)
    with open(CONFIG_FILE, encoding='utf-8') as f:
        return json.load(f)


def salvar_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def txt(elem, path, ns={'nfe': NS_NFE}):
    """Extrai texto de um subelemento xpath."""
    e = elem.find(path, ns)
    return (e.text or '').strip() if e is not None else ''


# ─────────────────────────────────────────────────────────────
# Certificado
# ─────────────────────────────────────────────────────────────
def extrair_pem(pfx_path, senha):
    """Lê o .pfx e retorna (cert_pem_bytes, key_pem_bytes)."""
    with open(pfx_path, 'rb') as f:
        pfx_data = f.read()
    pwd = senha.encode() if isinstance(senha, str) else senha
    priv_key, cert, _ = load_key_and_certificates(pfx_data, pwd)
    cert_pem = cert.public_bytes(Encoding.PEM)
    key_pem  = priv_key.private_bytes(Encoding.PEM, PrivateFormat.TraditionalOpenSSL, NoEncryption())
    return cert_pem, key_pem


# ─────────────────────────────────────────────────────────────
# Assinatura XML (XMLDSig — obrigatório pela SEFAZ)
# ─────────────────────────────────────────────────────────────
def assinar_xml(xml_element, cert_pem, key_pem):
    try:
        from signxml import XMLSigner, methods
        from cryptography.x509 import load_pem_x509_certificate
        from cryptography.hazmat.primitives.serialization import load_pem_private_key
    except ImportError:
        raise ImportError("Execute: pip install signxml")

    cert = load_pem_x509_certificate(cert_pem)
    key  = load_pem_private_key(key_pem, password=None)

    signer = XMLSigner(
        method=methods.enveloped,
        digest_algorithm='sha1',
        signature_algorithm='rsa-sha1',
        c14n_algorithm='http://www.w3.org/TR/2001/REC-xml-c14n-20010315',
    )
    return signer.sign(xml_element, key=key, cert=cert)


# ─────────────────────────────────────────────────────────────
# SOAP / SEFAZ
# ─────────────────────────────────────────────────────────────
def _construir_soap(cnpj, cuf, ambiente, ult_nsu, cert_pem, key_pem):
    dist_xml = etree.fromstring(
        f'<distDFeInt xmlns="{NS_NFE}" versao="1.01">'
        f'<tpAmb>{ambiente}</tpAmb>'
        f'<cUFAutor>{cuf}</cUFAutor>'
        f'<CNPJ>{cnpj}</CNPJ>'
        f'<distNSU><ultNSU>{ult_nsu}</ultNSU></distNSU>'
        f'</distDFeInt>'.encode()
    )
    signed     = assinar_xml(dist_xml, cert_pem, key_pem)
    signed_str = etree.tostring(signed, encoding='unicode')

    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
        f' xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
        f' xmlns:soap12="{SOAP_NS}">'
        f'<soap12:Header>'
        f'<nfeCabecMsg xmlns="{NS_DIST}">'
        f'<cUF>{cuf}</cUF><versaoDados>1.01</versaoDados>'
        f'</nfeCabecMsg>'
        f'</soap12:Header>'
        f'<soap12:Body>'
        f'<nfeDistDFeInteresse xmlns="{NS_DIST}">'
        f'<nfeDadosMsg>{signed_str}</nfeDadosMsg>'
        f'</nfeDistDFeInteresse>'
        f'</soap12:Body>'
        f'</soap12:Envelope>'
    )


def _post_sefaz(soap_body, cert_tmp, key_tmp):
    headers = {
        'Content-Type': 'application/soap+xml; charset=utf-8',
        'SOAPAction': '',
    }
    resp = requests.post(
        SEFAZ_URL,
        data=soap_body.encode('utf-8'),
        headers=headers,
        cert=(cert_tmp, key_tmp),
        verify=True,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.text


def _extrair_ret(resp_text):
    """Navega o envelope SOAP e retorna o elemento retDistDFeInt."""
    root = etree.fromstring(resp_text.encode('utf-8'))
    for elem in root.iter():
        if 'retDistDFeInt' in elem.tag:
            return elem
        if 'nfeDistDFeInteresseResult' in elem.tag and elem.text:
            try:
                return etree.fromstring(elem.text.encode('utf-8'))
            except Exception:
                return elem
    raise Exception("Resposta inesperada da SEFAZ — retDistDFeInt não encontrado.")


def _find(ret, tag):
    """Busca tag com ou sem namespace."""
    for ns in [NS_NFE, '']:
        full = f'{{{ns}}}{tag}' if ns else tag
        e = ret.find(f'.//{full}')
        if e is not None:
            return e
    return None


def consultar_pagina(cnpj, cuf, ambiente, ult_nsu, cert_pem, key_pem, cert_tmp, key_tmp):
    """Uma chamada ao SEFAZ. Retorna (documentos, ult_nsu_ret, max_nsu_ret)."""
    soap = _construir_soap(cnpj, cuf, ambiente, ult_nsu, cert_pem, key_pem)
    resp_text = _post_sefaz(soap, cert_tmp, key_tmp)
    ret = _extrair_ret(resp_text)

    cstat  = (_find(ret, 'cStat') or type('', (), {'text': '0'})()).text
    xmotiv = (_find(ret, 'xMotivo') or type('', (), {'text': ''})()).text

    # 137 = sem docs, 138 = docs encontrados
    if cstat not in ('137', '138'):
        raise Exception(f"SEFAZ cStat {cstat}: {xmotiv}")

    ult_nsu_ret = (_find(ret, 'ultNSU') or type('', (), {'text': ult_nsu})()).text
    max_nsu_ret = (_find(ret, 'maxNSU') or type('', (), {'text': ult_nsu_ret})()).text

    documentos = []
    for doc_zip in ret.iter(f'{{{NS_NFE}}}docZip'):
        nsu    = doc_zip.get('NSU', '')
        schema = doc_zip.get('schema', '')
        try:
            xml_bytes = gzip.decompress(base64.b64decode(doc_zip.text))
            documentos.append((nsu, schema, xml_bytes))
        except Exception as e:
            log(f"  ⚠ NSU {nsu}: falha ao descompactar — {e}")

    return documentos, ult_nsu_ret, max_nsu_ret


def baixar_todos(cnpj, cuf, ambiente, ult_nsu_inicial, cert_pem, key_pem):
    """Pagina pelos NSUs até trazer todos os documentos disponíveis."""
    # Arquivos temporários para o requests (mTLS)
    ct = tempfile.NamedTemporaryFile(delete=False, suffix='_cert.pem', mode='wb')
    kt = tempfile.NamedTemporaryFile(delete=False, suffix='_key.pem',  mode='wb')
    ct.write(cert_pem); kt.write(key_pem)
    ct.close(); kt.close()

    todos   = []
    cur_nsu = ult_nsu_inicial
    pagina  = 0

    try:
        while True:
            pagina += 1
            log(f"Consultando SEFAZ — página {pagina} | NSU: {cur_nsu}")
            docs, ult_ret, max_ret = consultar_pagina(
                cnpj, cuf, ambiente, cur_nsu, cert_pem, key_pem, ct.name, kt.name
            )
            todos.extend(docs)
            log(f"  → {len(docs)} docs recebidos | ultNSU: {ult_ret} | maxNSU: {max_ret}")

            if ult_ret == max_ret or ult_ret == cur_nsu:
                break
            cur_nsu = ult_ret
            time.sleep(1)  # respeitar intervalo mínimo da SEFAZ
    finally:
        os.unlink(ct.name)
        os.unlink(kt.name)

    return todos, cur_nsu


# ─────────────────────────────────────────────────────────────
# Parser NF-e XML
# ─────────────────────────────────────────────────────────────
def parsear_nfe(xml_bytes, cnpj_proprio):
    """
    Parseia a NF-e e retorna lista de dicts (um por item).
    Retorna [] se não for uma NF-e com itens (ex.: resNFe sumário).
    """
    try:
        root = etree.fromstring(xml_bytes)
    except Exception:
        return []

    ns = {'nfe': NS_NFE}

    # Suporta nfeProc (raiz) ou NFe diretamente
    nfe = root.find('.//nfe:NFe', ns)
    if nfe is None:
        nfe = root  # tenta como raiz

    inf = nfe.find('.//nfe:infNFe', ns)
    if inf is None:
        return []  # não é NF-e completa

    ide   = nfe.find('.//nfe:ide', ns)
    emit  = nfe.find('.//nfe:emit', ns)
    dest  = nfe.find('.//nfe:dest', ns)
    total = nfe.find('.//nfe:total/nfe:ICMSTot', ns)

    chave     = inf.get('Id', '').replace('NFe', '')
    emit_cnpj = txt(emit, 'nfe:CNPJ', ns) if emit is not None else ''
    dest_cnpj = txt(dest, 'nfe:CNPJ', ns) if dest is not None else ''
    tipo      = 'COMPRA' if dest_cnpj == cnpj_proprio else 'VENDA'

    dh_emi = txt(ide, 'nfe:dhEmi', ns) or txt(ide, 'nfe:dEmi', ns)
    data_emissao = dh_emi[:10] if dh_emi else ''

    cabecalho = {
        'chave_nfe':    chave,
        'tipo':         tipo,
        'data_emissao': data_emissao,
        'numero_nf':    txt(ide, 'nfe:nNF', ns),
        'serie':        txt(ide, 'nfe:serie', ns),
        'nat_op':       txt(ide, 'nfe:natOp', ns),
        'emit_cnpj':    emit_cnpj,
        'emit_nome':    txt(emit, 'nfe:xNome', ns)                 if emit is not None else '',
        'emit_uf':      txt(emit, 'nfe:enderEmit/nfe:UF', ns)      if emit is not None else '',
        'dest_cnpj':    dest_cnpj,
        'dest_nome':    txt(dest, 'nfe:xNome', ns)                 if dest is not None else '',
        'dest_uf':      txt(dest, 'nfe:enderDest/nfe:UF', ns)      if dest is not None else '',
        'vl_prod':      txt(total, 'nfe:vProd', ns)                if total is not None else '',
        'vl_frete':     txt(total, 'nfe:vFrete', ns)               if total is not None else '',
        'vl_desc_nf':   txt(total, 'nfe:vDesc', ns)                if total is not None else '',
        'vl_total':     txt(total, 'nfe:vNF', ns)                  if total is not None else '',
        'vl_icms':      txt(total, 'nfe:vICMS', ns)                if total is not None else '',
        'vl_pis':       txt(total, 'nfe:vPIS', ns)                 if total is not None else '',
        'vl_cofins':    txt(total, 'nfe:vCOFINS', ns)              if total is not None else '',
    }

    linhas = []
    for det in nfe.findall('.//nfe:det', ns):
        prod = det.find('nfe:prod', ns)
        if prod is None:
            continue
        linhas.append({**cabecalho,
            'num_item':   det.get('nItem', ''),
            'codigo':     txt(prod, 'nfe:cProd', ns),
            'ean':        txt(prod, 'nfe:cEAN', ns),
            'descricao':  txt(prod, 'nfe:xProd', ns),
            'ncm':        txt(prod, 'nfe:NCM', ns),
            'cfop':       txt(prod, 'nfe:CFOP', ns),
            'unidade':    txt(prod, 'nfe:uCom', ns),
            'quantidade': txt(prod, 'nfe:qCom', ns),
            'vl_unit':    txt(prod, 'nfe:vUnCom', ns),
            'vl_item':    txt(prod, 'nfe:vProd', ns),
            'desc_item':  txt(prod, 'nfe:vDesc', ns),
        })
    return linhas


def linha_para_row(l):
    return [
        l['chave_nfe'], l['data_emissao'], l['numero_nf'], l['serie'], l['nat_op'],
        l['emit_cnpj'], l['emit_nome'], l['emit_uf'],
        l['dest_cnpj'], l['dest_nome'], l['dest_uf'],
        l['num_item'], l['codigo'], l['ean'], l['descricao'],
        l['ncm'], l['cfop'], l['unidade'], l['quantidade'],
        l['vl_unit'], l['vl_item'], l['desc_item'],
        l['vl_prod'], l['vl_frete'], l['vl_desc_nf'], l['vl_total'],
        l['vl_icms'], l['vl_pis'], l['vl_cofins'],
    ]


# ─────────────────────────────────────────────────────────────
# Google Sheets
# ─────────────────────────────────────────────────────────────
def conectar_sheets(cred_path):
    creds = Credentials.from_service_account_file(cred_path, scopes=SHEETS_SCOPE)
    return gspread.authorize(creds)


def preparar_planilha(gc, nome):
    """Abre ou cria a planilha e garante as abas Compras/Vendas."""
    try:
        sh = gc.open(nome)
        log(f"Planilha '{nome}' encontrada.")
    except gspread.SpreadsheetNotFound:
        sh = gc.create(nome)
        log(f"Planilha '{nome}' criada: {sh.url}")

    for aba in ['Compras', 'Vendas']:
        try:
            ws = sh.worksheet(aba)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(aba, rows=1, cols=len(CABECALHO))
            ws.append_row(CABECALHO, value_input_option='RAW')
            # Formatar cabeçalho (negrito, fundo escuro)
            ws.format('A1:AC1', {
                'backgroundColor': {'red': 0.11, 'green': 0.17, 'blue': 0.28},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 0.22, 'green': 0.55, 'blue': 0.99}},
            })
            log(f"  Aba '{aba}' criada.")

    # Remove aba padrão "Planilha1" / "Sheet1" se existir
    for nome_padrao in ('Planilha1', 'Sheet1', 'Plan1'):
        try:
            sh.del_worksheet(sh.worksheet(nome_padrao))
        except Exception:
            pass

    return sh


def gravar_linhas(sh, compras, vendas):
    for aba_nome, linhas in [('Compras', compras), ('Vendas', vendas)]:
        if not linhas:
            log(f"  Nenhum item para '{aba_nome}'.")
            continue
        ws   = sh.worksheet(aba_nome)
        rows = [linha_para_row(l) for l in linhas]
        # Grava em lotes de 500 para evitar timeout
        for i in range(0, len(rows), 500):
            ws.append_rows(rows[i:i+500], value_input_option='USER_ENTERED')
        log(f"  ✅ '{aba_nome}': {len(rows)} itens gravados.")


# ─────────────────────────────────────────────────────────────
# Ponto de entrada
# ─────────────────────────────────────────────────────────────
def main():
    log("=" * 50)
    log("NF-e Downloader — SEFAZ DistribuicaoDFe")
    log("=" * 50)

    cfg      = carregar_config()
    cnpj     = re.sub(r'\D', '', cfg['cnpj'])
    cuf      = int(cfg.get('cuf', 35))
    ambiente = int(cfg.get('ambiente', 1))
    ult_nsu  = cfg.get('ultimo_nsu', '000000000000000').zfill(15)
    cred     = cfg.get('credentials_json', 'credentials.json')
    planilha = cfg.get('nome_planilha', 'NF-e SEFAZ')

    # 1. Certificado
    pfx_path = cfg['certificado_pfx']
    pfx_pwd  = cfg['senha_certificado']
    log(f"Carregando certificado: {pfx_path}")
    cert_pem, key_pem = extrair_pem(pfx_path, pfx_pwd)
    log("  ✅ Certificado OK.")

    # 2. Baixar da SEFAZ
    log(f"CNPJ: {cnpj} | cUF: {cuf} | Ambiente: {'Produção' if ambiente == 1 else 'Homologação'}")
    log(f"Último NSU: {ult_nsu}  (000...000 = baixar tudo desde o início)")
    todos, ult_nsu_final = baixar_todos(cnpj, cuf, ambiente, ult_nsu, cert_pem, key_pem)
    log(f"Total de documentos recebidos da SEFAZ: {len(todos)}")

    # 3. Parsear
    log("Parseando NF-e...")
    compras, vendas = [], []
    ignorados = 0
    for nsu, schema, xml_bytes in todos:
        # Só processa NF-e completas (procNFe). resNFe são resumos sem itens.
        if not any(s in schema for s in ('procNFe', 'nfeProc', 'NFe')):
            ignorados += 1
            continue
        for linha in parsear_nfe(xml_bytes, cnpj):
            (compras if linha['tipo'] == 'COMPRA' else vendas).append(linha)

    log(f"  Compras: {len(compras)} itens | Vendas: {len(vendas)} itens | Ignorados (resumos/outros): {ignorados}")

    if not compras and not vendas:
        log("\n⚠  Nenhum item encontrado.")
        log("   Se este é o primeiro uso, pode ser que as NF-e de compra ainda estejam")
        log("   como 'resNFe' (resumo). Nesse caso, faça a Manifestação do Destinatário")
        log("   no portal da SEFAZ para liberar o XML completo e rode o script novamente.")
        # Salva o NSU mesmo assim para não baixar tudo de novo
        cfg['ultimo_nsu'] = ult_nsu_final
        salvar_config(cfg)
        return

    # 4. Google Sheets
    log(f"Conectando ao Google Sheets (credencial: {cred})...")
    gc = conectar_sheets(cred)
    sh = preparar_planilha(gc, planilha)
    gravar_linhas(sh, compras, vendas)

    # 5. Salvar NSU para próxima execução
    cfg['ultimo_nsu'] = ult_nsu_final
    salvar_config(cfg)

    log(f"\nÚltimo NSU salvo: {ult_nsu_final}")
    log(f"Planilha: {sh.url}")
    log("✅ Concluído!")


if __name__ == '__main__':
    main()
