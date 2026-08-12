#!/usr/bin/env python3
"""
NF-e de Venda — recuperação pela SEFAZ
======================================

Contexto do problema
--------------------
O serviço NFeDistribuicaoDFe (NT 2014.002) distribui os DF-e *de interesse*
do CNPJ: aqueles em que ele é destinatário, transportador ou terceiro
referenciado. As NF-e que o próprio CNPJ **emite** não voltam por esse
endpoint — a SEFAZ parte do princípio de que o emitente já recebeu o
protocolo de autorização no momento da emissão e é obrigado a guardar o
XML por 5 anos.

Por isso o relatório mostra "NF-e Vendas: 0" mesmo com 46 mil documentos
baixados: o ramo VENDA de `parsear_nfe` nunca é alcançado.

O que ainda dá para extrair da SEFAZ
------------------------------------
Eventos. O emitente RECEBE os eventos das próprias notas (manifestação do
destinatário, cancelamento, carta de correção). Cada evento carrega a
`chNFe`, e a chave de acesso é auto-descritiva:

    posições (0-based)   conteúdo
     0..2                cUF
     2..6                AAMM da emissão
     6..20               CNPJ do emitente     ← filtro das suas vendas
    20..22               modelo (55 = NF-e)
    22..25               série
    25..34               número da NF
    34..35               tpEmis
    35..43               código numérico
    43..44               DV

Ou seja: dos eventos sai chave, número, série e mês de emissão das suas
vendas — sem detalhe de item. Cobertura é parcial por natureza: só aparece
a venda que recebeu algum evento.

Modos
-----
    py vendas.py chave <44 digitos>   testa consChNFe numa chave sua
    py vendas.py dump                 rebaixa persistindo XML cru (resumível)
    py vendas.py scan                 offline: varre xmls/ e monta as chaves
    py vendas.py buscar               consChNFe em lote sobre as chaves
"""

import json
import os
import re
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

from lxml import etree

from nfe_downloader import (
    NS_NFE, XMLS_DIR,
    carregar_config, salvar_config, extrair_pem, log,
    classificar_schema, salvar_xml,
    _post_sefaz, _extrair_ret, _find,
    parsear_nfe,
)

ESTADO_DUMP   = 'dump_estado.json'
CHAVES_VENDAS = 'chaves_vendas.json'

# tpEvento mais comuns na NT 2014.002
TP_EVENTO = {
    '110110': 'Carta de Correção',
    '110111': 'Cancelamento',
    '110112': 'Cancelamento por substituição',
    '110140': 'EPEC',
    '111500': 'Pedido de prorrogação',
    '111501': 'Pedido de prorrogação 2',
    '210200': 'Confirmação da Operação',
    '210210': 'Ciência da Operação',
    '210220': 'Desconhecimento da Operação',
    '210240': 'Operação não Realizada',
}


# ─────────────────────────────────────────────────────────────
# Chave de acesso
# ─────────────────────────────────────────────────────────────
def partes_chave(chave):
    """Decompõe a chave de 44 dígitos nos campos que ela codifica."""
    c = re.sub(r'\D', '', chave or '')
    if len(c) != 44:
        return None
    return {
        'chave':     c,
        'cuf':       c[0:2],
        'aamm':      c[2:6],
        'competencia': f'20{c[2:4]}-{c[4:6]}',
        'cnpj_emit': c[6:20],
        'modelo':    c[20:22],
        'serie':     str(int(c[22:25])),
        'numero':    str(int(c[25:34])),
        'tp_emis':   c[34:35],
        'dv':        c[43:44],
    }


def chave_valida(chave):
    """Valida o DV (módulo 11) da chave de acesso."""
    c = re.sub(r'\D', '', chave or '')
    if len(c) != 44:
        return False
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    soma = sum(int(d) * pesos[i % 8] for i, d in enumerate(reversed(c[:43])))
    resto = soma % 11
    dv = 0 if resto in (0, 1) else 11 - resto
    return dv == int(c[43])


# ─────────────────────────────────────────────────────────────
# SOAP — consulta por chave (consChNFe)
# ─────────────────────────────────────────────────────────────
NS_DIST = 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe'
SOAP_NS = 'http://www.w3.org/2003/05/soap-envelope'


def _soap_cons_chave(cnpj, cuf, ambiente, chave):
    """Mesmo envelope do distDFeInt, trocando distNSU por consChNFe."""
    xml_body = (
        f'<distDFeInt xmlns="{NS_NFE}" versao="1.01">'
        f'<tpAmb>{ambiente}</tpAmb>'
        f'<cUFAutor>{cuf}</cUFAutor>'
        f'<CNPJ>{cnpj}</CNPJ>'
        f'<consChNFe><chNFe>{chave}</chNFe></consChNFe>'
        f'</distDFeInt>'
    )
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
        f'<nfeDadosMsg>{xml_body}</nfeDadosMsg>'
        f'</nfeDistDFeInteresse>'
        f'</soap12:Body>'
        f'</soap12:Envelope>'
    )


class _Cert:
    """Materializa cert/key PEM em arquivos temporários para o mTLS."""

    def __init__(self, cert_pem, key_pem):
        self.cert_pem, self.key_pem = cert_pem, key_pem

    def __enter__(self):
        self.ct = tempfile.NamedTemporaryFile(delete=False, suffix='_cert.pem', mode='wb')
        self.kt = tempfile.NamedTemporaryFile(delete=False, suffix='_key.pem', mode='wb')
        self.ct.write(self.cert_pem); self.kt.write(self.key_pem)
        self.ct.close(); self.kt.close()
        return self.ct.name, self.kt.name

    def __exit__(self, *exc):
        for f in (self.ct.name, self.kt.name):
            try:
                os.unlink(f)
            except OSError:
                pass


def consultar_chave(cnpj, cuf, ambiente, chave, cert_tmp, key_tmp):
    """
    Consulta uma chave via consChNFe.

    Retorna (cstat, xmotivo, [(nsu, schema, xml_bytes), ...]).
    """
    import base64
    import gzip

    soap = _soap_cons_chave(cnpj, cuf, ambiente, chave)
    ret = _extrair_ret(_post_sefaz(soap, cert_tmp, key_tmp))

    e_cstat = _find(ret, 'cStat')
    e_xmot = _find(ret, 'xMotivo')
    cstat = e_cstat.text if e_cstat is not None else '0'
    xmotiv = e_xmot.text if e_xmot is not None else ''

    docs = []
    for doc_zip in ret.iter(f'{{{NS_NFE}}}docZip'):
        nsu = doc_zip.get('NSU', '')
        schema = doc_zip.get('schema', '')
        try:
            docs.append((nsu, schema, gzip.decompress(base64.b64decode(doc_zip.text))))
        except Exception as e:
            log(f"  ⚠ NSU {nsu}: falha ao descompactar — {e}")
    return cstat, xmotiv, docs


def _abrir_certificado(cfg):
    cert_pem, key_pem = extrair_pem(cfg['certificado_pfx'], cfg['senha_certificado'])
    return cert_pem, key_pem


# ─────────────────────────────────────────────────────────────
# Modo: chave
# ─────────────────────────────────────────────────────────────
def modo_chave(argv):
    if not argv:
        print("Uso: py vendas.py chave <44 digitos>")
        return 1

    chave = re.sub(r'\D', '', argv[0])
    p = partes_chave(chave)
    if not p:
        print(f"❌ Chave precisa ter 44 dígitos (recebi {len(chave)}).")
        return 1

    cfg = carregar_config()
    cnpj = re.sub(r'\D', '', cfg['cnpj'])
    cuf = int(cfg.get('cuf', 35))
    ambiente = int(cfg.get('ambiente', 1))

    log("=" * 60)
    log("Teste consChNFe — a SEFAZ devolve o XML de uma venda própria?")
    log("=" * 60)
    log(f"Chave        : {chave}")
    log(f"DV válido    : {'sim' if chave_valida(chave) else 'NÃO — confira a chave'}")
    log(f"CNPJ emitente: {p['cnpj_emit']}  (seu CNPJ: {cnpj})")
    log(f"Competência  : {p['competencia']} | Série {p['serie']} | Nº {p['numero']}")
    if p['cnpj_emit'] != cnpj:
        log("⚠  Esta chave NÃO foi emitida pelo seu CNPJ — o teste perde o sentido.")

    cert_pem, key_pem = _abrir_certificado(cfg)
    with _Cert(cert_pem, key_pem) as (ct, kt):
        cstat, xmotiv, docs = consultar_chave(cnpj, cuf, ambiente, chave, ct, kt)

    log(f"→ cStat {cstat}: {xmotiv}")
    if not docs:
        log("Nenhum documento devolvido.")
        if cstat == '656':
            log("Consumo Indevido — o CNPJ está bloqueado por ~1 hora.")
        return 0

    for nsu, schema, xml_bytes in docs:
        destino = salvar_xml(nsu, schema, xml_bytes)
        tipo = classificar_schema(schema)
        log(f"  NSU {nsu} | schema {schema} | tipo {tipo} → {destino}")
        if tipo == 'nfe':
            linhas = parsear_nfe(xml_bytes, cnpj)
            log(f"  ✅ NF-e COMPLETA — {len(linhas)} itens.")
            if linhas:
                log(f"     Emitente: {linhas[0]['emit_nome']} | Dest.: {linhas[0]['dest_nome']}")
                log(f"     Total: R$ {linhas[0]['vl_total']} | Classificado como: {linhas[0]['tipo']}")
                log("     → consChNFe resolve o problema: use 'py vendas.py buscar'.")
        elif tipo == 'resNFe':
            log("  ⚠ Só o RESUMO (resNFe) — sem itens. consChNFe não basta.")
    return 0


# ─────────────────────────────────────────────────────────────
# Modo: dump
# ─────────────────────────────────────────────────────────────
def _carregar_estado():
    if Path(ESTADO_DUMP).exists():
        with open(ESTADO_DUMP, encoding='utf-8') as f:
            return json.load(f)
    return {}


def _salvar_estado(estado):
    with open(ESTADO_DUMP, 'w', encoding='utf-8') as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def modo_dump(argv):
    """
    Rebaixa os documentos persistindo o XML cru, com progresso salvo a cada
    página. Por padrão retoma do último NSU conhecido — NUNCA do zero, porque
    reiniciar do NSU 0 é tratado como Consumo Indevido (cStat 656) e bloqueia
    o CNPJ por 1 hora.
    """
    from nfe_downloader import consultar_pagina

    do_zero = '--zero' in argv
    cfg = carregar_config()
    cnpj = re.sub(r'\D', '', cfg['cnpj'])
    cuf = int(cfg.get('cuf', 35))
    ambiente = int(cfg.get('ambiente', 1))

    estado = _carregar_estado()
    if do_zero:
        log("⚠  --zero: reiniciando do NSU 0. Se a SEFAZ responder 656, o CNPJ")
        log("   fica bloqueado por ~1 hora. Só use se o histórico foi perdido.")
        cur_nsu = '000000000000000'
    else:
        cur_nsu = estado.get('ultimo_nsu') or cfg.get('ultimo_nsu', '000000000000000')
    cur_nsu = str(cur_nsu).zfill(15)

    log(f"Dump a partir do NSU {cur_nsu} → '{XMLS_DIR}/'")
    cert_pem, key_pem = _abrir_certificado(cfg)

    total = 0
    pagina = 0
    contagem = Counter()

    with _Cert(cert_pem, key_pem) as (ct, kt):
        while True:
            pagina += 1
            try:
                docs, ult_ret, max_ret = consultar_pagina(
                    cnpj, cuf, ambiente, cur_nsu, cert_pem, key_pem, ct, kt
                )
            except Exception as e:
                log(f"❌ Página {pagina} falhou: {e}")
                if '656' in str(e):
                    log("   Consumo Indevido — aguarde ~1 hora e rode de novo.")
                    log(f"   Progresso preservado: NSU {cur_nsu}.")
                break

            for nsu, schema, xml_bytes in docs:
                salvar_xml(nsu, schema, xml_bytes)
                contagem[classificar_schema(schema)] += 1
            total += len(docs)

            log(f"  pág {pagina} | +{len(docs)} docs | total {total} "
                f"| ultNSU {ult_ret} | maxNSU {max_ret}")

            estado['ultimo_nsu'] = ult_ret
            estado['total_baixado'] = total
            _salvar_estado(estado)

            if ult_ret == max_ret or ult_ret == cur_nsu:
                log("Fim da paginação.")
                break
            cur_nsu = ult_ret
            time.sleep(1)

    if contagem:
        log("Distribuição: " + ' | '.join(f'{k}: {v}' for k, v in sorted(contagem.items())))
    log(f"✅ {total} documentos salvos em '{XMLS_DIR}/'. Rode: py vendas.py scan")
    return 0


# ─────────────────────────────────────────────────────────────
# Modo: scan (offline)
# ─────────────────────────────────────────────────────────────
def _texto_por_sufixo(root, sufixo):
    """Primeiro elemento cujo tag termina em `sufixo`, ignorando namespace."""
    for elem in root.iter():
        tag = elem.tag if isinstance(elem.tag, str) else ''
        if tag.rsplit('}', 1)[-1] == sufixo and elem.text:
            return elem.text.strip()
    return ''


def modo_scan(argv):
    """
    Varre `xmls/` sem tocar na SEFAZ.

    - Eventos com chNFe cujo CNPJ emitente é o seu  → suas VENDAS
    - resNFe                                        → COMPRAS não manifestadas
    """
    base = Path(XMLS_DIR)
    if not base.exists():
        log(f"❌ Diretório '{XMLS_DIR}/' não existe. Rode 'py vendas.py dump' antes.")
        return 1

    cfg = carregar_config()
    cnpj = re.sub(r'\D', '', cfg['cnpj'])

    arquivos = sorted(base.glob('*.xml'))
    log(f"Varrendo {len(arquivos)} arquivos em '{XMLS_DIR}/' (CNPJ {cnpj})...")

    vendas = {}          # chave → dados agregados
    compras_resumo = {}  # chave → resNFe
    eventos_terceiros = 0
    tipos = Counter()
    ilegiveis = 0

    for arq in arquivos:
        try:
            root = etree.fromstring(arq.read_bytes())
        except Exception:
            ilegiveis += 1
            continue

        # salvar_xml grava como '{nsu}_{tipo}.xml' — o tipo já vem classificado.
        tipo = arq.stem.split('_', 1)[1] if '_' in arq.stem else 'outro'
        tipos[tipo] += 1

        if tipo == 'evento':
            chave = _texto_por_sufixo(root, 'chNFe')
            p = partes_chave(chave)
            if not p:
                continue
            if p['cnpj_emit'] != cnpj:
                eventos_terceiros += 1
                continue

            tp = _texto_por_sufixo(root, 'tpEvento')
            reg = vendas.setdefault(chave, {
                **p,
                'eventos': [],
                'cancelada': False,
            })
            reg['eventos'].append({
                'tpEvento': tp,
                'descricao': TP_EVENTO.get(tp, _texto_por_sufixo(root, 'xEvento') or tp),
                'dhEvento': _texto_por_sufixo(root, 'dhEvento'),
            })
            if tp in ('110111', '110112'):
                reg['cancelada'] = True

        elif tipo == 'resNFe':
            chave = _texto_por_sufixo(root, 'chNFe')
            p = partes_chave(chave)
            if not p:
                continue
            compras_resumo[chave] = {
                **p,
                'emit_nome': _texto_por_sufixo(root, 'xNome'),
                'dh_emi': _texto_por_sufixo(root, 'dhEmi'),
                'vl_nf': _texto_por_sufixo(root, 'vNF'),
                'situacao': _texto_por_sufixo(root, 'cSitNFe'),
            }

    log("")
    log("Tipos de documento em disco: " + (
        ' | '.join(f'{k}: {v}' for k, v in sorted(tipos.items())) or '(nenhum)'))
    if ilegiveis:
        log(f"  {ilegiveis} arquivos ilegíveis (ignorados).")
    log("")
    log(f"✅ VENDAS detectadas via eventos : {len(vendas)} chaves")
    log(f"   (eventos de notas de terceiros: {eventos_terceiros})")
    log(f"✅ COMPRAS em resumo (resNFe)    : {len(compras_resumo)} chaves")
    log("   resNFe = compra ainda não manifestada; o XML completo só é")
    log("   liberado após a Manifestação do Destinatário.")

    if vendas:
        canceladas = sum(1 for v in vendas.values() if v['cancelada'])
        comps = Counter(v['competencia'] for v in vendas.values())
        log("")
        log(f"   Canceladas: {canceladas} | Ativas: {len(vendas) - canceladas}")
        log("   Por competência:")
        for comp, qtd in sorted(comps.items()):
            log(f"     {comp}: {qtd}")

    saida = {
        'cnpj': cnpj,
        'vendas': sorted(vendas.values(), key=lambda v: (v['competencia'], int(v['numero']))),
        'compras_resumo': sorted(compras_resumo.values(),
                                 key=lambda v: (v['competencia'], int(v['numero']))),
    }
    with open(CHAVES_VENDAS, 'w', encoding='utf-8') as f:
        json.dump(saida, f, ensure_ascii=False, indent=2)
    log("")
    log(f"→ Gravado em '{CHAVES_VENDAS}'. Para tentar o XML completo:")
    log("  py vendas.py buscar")
    return 0


# ─────────────────────────────────────────────────────────────
# Modo: buscar
# ─────────────────────────────────────────────────────────────
def modo_buscar(argv):
    """consChNFe em lote sobre as chaves detectadas pelo scan."""
    if not Path(CHAVES_VENDAS).exists():
        log(f"❌ '{CHAVES_VENDAS}' não existe. Rode 'py vendas.py scan' antes.")
        return 1

    intervalo = 2.0
    limite = None
    for i, a in enumerate(argv):
        if a == '--intervalo' and i + 1 < len(argv):
            intervalo = float(argv[i + 1])
        if a == '--limite' and i + 1 < len(argv):
            limite = int(argv[i + 1])

    with open(CHAVES_VENDAS, encoding='utf-8') as f:
        dados = json.load(f)
    chaves = [v['chave'] for v in dados['vendas'] if not v.get('cancelada')]
    if limite:
        chaves = chaves[:limite]

    if not chaves:
        log("Nenhuma chave de venda ativa para consultar.")
        return 0

    cfg = carregar_config()
    cnpj = re.sub(r'\D', '', cfg['cnpj'])
    cuf = int(cfg.get('cuf', 35))
    ambiente = int(cfg.get('ambiente', 1))

    log(f"Consultando {len(chaves)} chaves via consChNFe (intervalo {intervalo}s)...")
    cert_pem, key_pem = _abrir_certificado(cfg)

    stats = Counter()
    completas = 0

    with _Cert(cert_pem, key_pem) as (ct, kt):
        for i, chave in enumerate(chaves, 1):
            try:
                cstat, xmotiv, docs = consultar_chave(cnpj, cuf, ambiente, chave, ct, kt)
            except Exception as e:
                log(f"  [{i}/{len(chaves)}] {chave} — erro: {e}")
                stats['erro'] += 1
                time.sleep(intervalo)
                continue

            stats[cstat] += 1
            tipos_doc = []
            for nsu, schema, xml_bytes in docs:
                salvar_xml(nsu, schema, xml_bytes)
                t = classificar_schema(schema)
                tipos_doc.append(t)
                if t == 'nfe':
                    completas += 1

            log(f"  [{i}/{len(chaves)}] {chave} → cStat {cstat} "
                f"{'/'.join(tipos_doc) if tipos_doc else '(sem doc)'}")

            if cstat == '656':
                log("❌ Consumo Indevido — interrompendo. Aguarde ~1 hora.")
                break
            time.sleep(intervalo)

    log("")
    log("Resumo por cStat: " + ' | '.join(f'{k}: {v}' for k, v in stats.most_common()))
    log(f"NF-e completas recuperadas: {completas}")
    if completas:
        log("→ Rode 'py nfe_downloader.py' para gravar os itens na planilha.")
    else:
        log("→ Nenhuma NF-e completa. A SEFAZ não devolve XML de venda própria;")
        log("  o XML tem de vir do sistema emissor.")
    return 0


# ─────────────────────────────────────────────────────────────
MODOS = {
    'chave':  modo_chave,
    'dump':   modo_dump,
    'scan':   modo_scan,
    'buscar': modo_buscar,
}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in MODOS:
        print(__doc__)
        return 1
    return MODOS[sys.argv[1]](sys.argv[2:])


if __name__ == '__main__':
    sys.exit(main())
