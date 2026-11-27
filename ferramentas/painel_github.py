"""Painel "Performance de vendas marca Dellamed" no GitHub (biancalimads/Resultados-Dellamed).

Atualização diária feita pela tarefa MASTER DELLAMED, depois de gerar o Excel do MASTER:
  python3 painel_github.py atualizar INDEX.html MASTER.xlsx REF HOJE PASTA_SAIDA PENDENTES [RESUMO_MES_ANTERIOR.json]
    - converte o Excel do MASTER nas linhas do painel (mesma regra do leitor de planilha do painel);
    - tranca (cifra) as linhas do mês com a chave pública que está no próprio index.html e troca o bloco desse mês;
      os outros meses não são tocados (a tarefa nunca precisa do PIN);
    - atualiza MES_ATUAL_LABEL e o carimbo data-saved;
    - grava PASTA_SAIDA/resumo_AAAA-MM.json (totais por dia do mês, com somas de conferência), que a tarefa
      guarda no projeto em claude/painel-github/resumo/AAAA-MM.json: é a base do mês anterior no mês seguinte;
    - grava PASTA_SAIDA/email.json com assunto, html e texto do e-mail de resumo (modelo do Monitor + link do painel).

Outros comandos:
  python3 painel_github.py chaves PIN SAIDA.json          gera o par de chaves (a privada sai trancada com o PIN)
  python3 painel_github.py montar INDEX.html CHAVES.json LINHAS.json   grava todos os meses (uso único na implantação)
  python3 painel_github.py conferir INDEX.html PIN        destranca com o PIN e mostra linhas e valor por mês (teste)

Formato no index.html:
  const EMBEDDED_KEYS = '{"pub":..., "salt":..., "iter":..., "iv":..., "priv":...}';
  const EMBEDDED_MONTHS = '{"AAAA-MM-01": "<bloco>", ...}';
  bloco = base64( chave AES do bloco cifrada com RSA-OAEP/SHA-256 (384 bytes) + iv (12) + AES-GCM(gzip(JSON das linhas)) )
  privada = AES-GCM(PBKDF2-SHA256(PIN, salt, iter), PKCS8)
"""
import sys, os, re, json, gzip, base64, datetime, collections, unicodedata

LINK = 'https://biancalimads.github.io/Resultados-Dellamed/'
ITER = 600000

# ---------------------------------------------------------------- conversão do MASTER
ITEM_CANON = {"CADEIRA DE RODAS D800": "CADEIRA MOTORIZADA D800", "CADEIRA DE RODAS D900": "CADEIRA MOTORIZADA D900", "CADEIRA DE RODAS D750": "CADEIRA MOTORIZADA D750", "CADEIRA DE RODAS D1000": "CADEIRA MOTORIZADA D1000", "CADEIRA DE RODAS OPTIMUS": "CADEIRA MOTORIZADA OPTIMUSMED", "SPACEMED": "CADEIRA MOTORIZADA SPACEMED", "CADEIRA DE RODAS D600 POP": "CADEIRA DE RODAS D600 POP T44", "BOTA PNEUMATICA": "BOTA PNEUMATICA AIR RECOVERY", "CAMA 2 MOVIMENTOS": "CAMA 2 MOVIMENTOS DX2", "CAMA 3 MOVIMENTOS": "CAMA 3 MOVIMENTOS DX3", "PEÇA": "PEÇA ACESSÓRIO", "UMIDIFICADOR UMILIGHT": "UMIDIFICADOR UMI LIGHT", "MONITOR DE BRAÇO COM VOZ ULTRA": "MONITOR DE BRACO COM VOZ ULTRA", "MASSAGEADOR SHIATSU PÉS": "MASSAGEADOR SHIATSU PES", "BALANÇA DIGITAL": "BALANCA DIGITAL", "BALANÇA DE BIOIMPEDANCIA": "BALANCA DE BIOIMPEDANCIA", "COLETE MAX PUTTI BAIXO": "COLETE MAX PUTTI CINZA ESCURO BAIXO", "COLETE MAX PUTTI ALTO": "COLETE MAX PUTTI CINZA ESCURO ALTO", "COLETE PUTTI BAIXO": "COLETE PUTTI COR NORMAL BAIXO", "MACA": "MACA MALETA", "JOELHEIRA ARTICULADA CINTAS CRUZADAS": "JOELHEIRA ARTICULADA COM CINTAS CRUZADAS", "JOELHEIRA ORIFÍCIO + REFORÇO PATELAR": "JOELHEIRACOM ORIFICIO E REFORCO PATELAR", "IMOBILIZADOR DE JOELHO": "IMOBILIZADOR", "IMOBILIZADOR DE JOELHO LARGO": "IMOBILIZADOR", "IMOBILIZADOR DE JOELHO ESTREITO": "IMOBILIZADOR", "ORTESE BILATERAL DE PUNHO": "TALA LONGA BILATERAL DE PUNHO", "ESPAÇADOR": "ESPACADOR DELLAMED", "TENNIS ELBOW": "TENIS ELBOW", "BOLSA DE AGUA QUENTE": "BOLSA AGUA QUENTE", "TALA LONGA COM POLEGAR": "ORTESE LONGA COM POLEGAR", "JOELHEIRA COM ORIFICIO E REFORCO PATELAR": "JOELHEIRACOM ORIFICIO E REFORCO PATELAR"}

def _sa(s): return ''.join(c for c in unicodedata.normalize('NFD', s) if not unicodedata.combining(c))

def _canon(v):
    s = str(v if v is not None else '').strip() or '-'
    if _sa(s).upper() == 'PNEUMATICO': s = 'PNEUMÁTICO'
    return ITEM_CANON.get(s, s)

def _txt(v): return '' if v is None else str(v).strip()

def _cat(v):
    t = ' '.join(_txt(v).split())
    return '' if t in ('', '-', 'Não informado') else t

def converter(path):
    """Excel do MASTER -> [Vendedor, Mês, Cat02, Modelo, Item, Valor, Qtd, PM, Data, Pneu, Cat03, Cat04, Cat05, Título]"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True); out = []
    for ws in wb.worksheets:
        is_equip = _sa(ws.title).upper().replace(' ', '') == 'EQUIPAMENTOMEDICO'
        it = ws.iter_rows(values_only=True); next(it)
        for r in it:
            if r[0] is None and r[2] is None: continue
            if _txt(r[15]) != 'DELLAMED': continue
            d = r[13].strftime('%Y-%m-%d')
            itv = _canon(r[14]); c5 = _txt(r[6])
            pn = 1 if is_equip and (itv == 'PNEUMÁTICO' or (c5 == 'Colchões Anti Escaras' and itv != '-' and itv.lower() not in ('não informado', 'nao informado'))) else 0
            q = r[8] or 0; pm = r[9] or 0
            out.append([_txt(r[1]) or '-', d[:7] + '-01', _txt(r[3]) or '-', r[11] or 'Não informado', itv,
                        float(r[7] or 0), int(round(q)), round(float(pm) * 100) / 100, d, pn,
                        _cat(r[4]), _cat(r[5]), _cat(r[6]), ' '.join(_txt(r[2]).split())])
    return out

# ---------------------------------------------------------------- criptografia
def _b64(b): return base64.b64encode(b).decode()
def _unb64(s): return base64.b64decode(s)

def gerar_chaves(pin):
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.primitives import serialization, hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    priv = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    pub = priv.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
    pk8 = priv.private_bytes(serialization.Encoding.DER, serialization.PrivateFormat.PKCS8, serialization.NoEncryption())
    salt, iv = os.urandom(16), os.urandom(12)
    kek = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pin.encode())
    return {'pub': _b64(pub), 'salt': _b64(salt), 'iter': ITER, 'iv': _b64(iv), 'priv': _b64(AESGCM(kek).encrypt(iv, pk8, None))}

def trancar(pub_b64, rows):
    from cryptography.hazmat.primitives import serialization, hashes
    from cryptography.hazmat.primitives.asymmetric import padding
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    pub = serialization.load_der_public_key(_unb64(pub_b64))
    payload = gzip.compress(json.dumps(rows, ensure_ascii=False, separators=(',', ':')).encode('utf-8'), 9)
    k, iv = AESGCM.generate_key(256), os.urandom(12)
    wk = pub.encrypt(k, padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None))
    assert len(wk) == 384
    return _b64(wk + iv + AESGCM(k).encrypt(iv, payload, None))

def abrir_privada(keys, pin):
    from cryptography.hazmat.primitives import serialization, hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    kek = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=_unb64(keys['salt']), iterations=keys['iter']).derive(pin.encode())
    return serialization.load_der_private_key(AESGCM(kek).decrypt(_unb64(keys['iv']), _unb64(keys['priv']), None), None)

def destrancar(priv, bloco):
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import padding
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    u = _unb64(bloco)
    k = priv.decrypt(u[:384], padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None))
    return json.loads(gzip.decompress(AESGCM(k).decrypt(u[384:396], u[396:], None)))

# ---------------------------------------------------------------- index.html
def _ler_const(html, nome):
    m = re.search(r"const %s = '([^']*)';" % nome, html)
    if not m: raise SystemExit('ERRO: o index.html não tem a constante %s (versão antiga do painel?).' % nome)
    return m.group(1)

def _troca_const(html, nome, valor):
    assert "'" not in valor
    novo, n = re.subn(r"const %s = '[^']*';" % nome, lambda m: "const %s = '%s';" % (nome, valor), html, count=1)
    if n != 1: raise SystemExit('ERRO: não achei a constante %s no index.html.' % nome)
    return novo

def _carimbo(html):
    agora = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
    novo, n = re.subn(r'data-saved="[^"]*"', 'data-saved="%s"' % agora, html, count=1)
    return novo

def gravar_meses(html, keys, por_mes):
    meses = json.loads(_ler_const(html, 'EMBEDDED_MONTHS'))
    for m, rows in por_mes.items():
        meses[m] = trancar(keys['pub'], rows)
    meses = dict(sorted(meses.items()))
    html = _troca_const(html, 'EMBEDDED_MONTHS', json.dumps(meses, separators=(',', ':')))
    html = _troca_const(html, 'MES_ATUAL_LABEL', max(meses))
    return _carimbo(html)

# ---------------------------------------------------------------- e-mail de resumo
SEM = ["segunda", "terça", "quarta", "quinta", "sexta", "sábado", "domingo"]
SEMC = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
def br(d): return d.strftime('%d/%m/%Y')
def fmt0(v): return "{:,.0f}".format(v).replace(",", ".")
def fmt2(v): return "{:,.2f}".format(v).replace(",", "X").replace(".", ",").replace("X", ".")
def fpct(v): return "sem base" if v is None else ("+" if v > 0 else "") + ("%.1f" % v).replace(".", ",") + "%"
def var(a, b): return None if not b else (a / b - 1) * 100

def resumo_mes(rows):
    acc = collections.defaultdict(lambda: [0.0, 0])
    for r in rows:
        acc[r[8]][0] += r[5]; acc[r[8]][1] += r[6]
    dias = {d: [round(g, 2), q] for d, (g, q) in sorted(acc.items())}
    return {'mes': rows[0][1][:7], 'dias': dias, 'soma_valor': round(sum(v[0] for v in dias.values()), 2),
            'soma_pecas': sum(v[1] for v in dias.values())}

def ler_resumo_anterior(path):
    """Lê o resumo do mês anterior (copiado do projeto) e confere as somas: se algo foi copiado errado, para."""
    if not path or not os.path.exists(path): return {}
    r = json.load(open(path, encoding='utf-8'))
    sv = round(sum(v[0] for v in r['dias'].values()), 2); sp = sum(v[1] for v in r['dias'].values())
    if abs(sv - r['soma_valor']) > 0.05 or sp != r['soma_pecas']:
        raise SystemExit('ERRO: o resumo do mês anterior não confere (soma %.2f x %.2f, peças %s x %s). Copie de novo o doc do projeto.' % (sv, r['soma_valor'], sp, r['soma_pecas']))
    return r['dias']

def atualizar_resumo(resumo, rows):
    """resumo = {"dias": {"AAAA-MM-DD": [valor, peças]}}; troca todos os dias do(s) mês(es) das linhas."""
    dias = resumo.setdefault('dias', {})
    meses = {r[1][:7] for r in rows}
    for d in [d for d in dias if d[:7] in meses]: del dias[d]
    acc = collections.defaultdict(lambda: [0.0, 0])
    for r in rows:
        acc[r[8]][0] += r[5]; acc[r[8]][1] += r[6]
    for d, (g, q) in acc.items(): dias[d] = [round(g, 2), q]
    resumo['dias'] = dict(sorted(dias.items()))
    return resumo

def montar_email(rows, resumo, ref, pendentes=0):
    dias = resumo['dias']
    def tot(a, b):
        g = q = 0.0; d = a
        while d <= b:
            v = dias.get(d.isoformat())
            if v: g += v[0]; q += v[1]
            d += datetime.timedelta(days=1)
        return g, q
    m0 = ref.replace(day=1)
    pm_fim = m0 - datetime.timedelta(days=1); pm0 = pm_fim.replace(day=1); pref = pm0.replace(day=min(ref.day, pm_fim.day))
    gd, qd = tot(ref, ref); g7, q7 = tot(ref - datetime.timedelta(days=6), ref)
    gp, qp = tot(pref, pref); gm, qm = tot(m0, ref); gpm, qpm = tot(pm0, pref)
    mes = [r for r in rows if m0.isoformat() <= r[8] <= ref.isoformat()]
    vm = len({r[0] for r in mes}); am = len({(r[0], r[13]) for r in mes if r[13]})
    tm, tm7 = (gd / qd if qd else 0), (g7 / q7 if q7 else 0)
    cor = lambda v: "#6b7383" if v is None else ("#1b7f4b" if v > 0.05 else ("#a6192e" if v < -0.05 else "#6b7383"))
    pd = "%s (%s)" % (pref.strftime("%d/%m"), SEMC[pref.weekday()])
    per_ant = "%s a %s" % (pm0.strftime("%d/%m"), pref.strftime("%d/%m"))
    v1, v2, v3 = var(gd, g7 / 7), var(qd, q7 / 7), var(tm, tm7)
    vp, vmm = var(gd, gp), var(gm, gpm)
    linhas = [
        ("Faturamento do dia", "R$ " + fmt0(gd), fpct(v1) + " vs. média de 7 dias", cor(v1)),
        ("", "", fpct(vp) + " vs. " + pd + ", mesmo dia do mês anterior", cor(vp)),
        ("Peças do dia", fmt0(qd), fpct(v2) + " vs. média de 7 dias", cor(v2)),
        ("Ticket médio", "R$ " + fmt2(tm), fpct(v3) + " vs. média de 7 dias", cor(v3)),
        ("Faturamento do mês (%s a %s)" % (m0.strftime("%d/%m"), ref.strftime("%d/%m")), "R$ " + fmt0(gm),
         fpct(vmm) + " vs. " + per_ant + " (R$ " + fmt0(gpm) + ")", cor(vmm)),
        ("", "", "%s vendedores · %s anúncios no mês" % (fmt0(vm), fmt0(am)), "#6b7383")]
    sub = "Resultados Dellamed · ML · dados de %s" % br(ref)
    html = ('<div style="font-family:Arial,sans-serif;color:#282727;font-size:14px;line-height:1.5">'
            '<p style="margin:0 0 4px;color:#1e2c4f;font-size:18px"><b>Resultados de Vendas</b> · Mercado Livre · marca Dellamed</p>'
            '<p style="margin:0 0 16px;color:#6b7383">Dados de %s (%s) · fonte: app Nubimetrics/Qlik Cloud</p>'
            '<table cellpadding="6" style="border-collapse:collapse;font-size:14px">' % (br(ref), SEM[ref.weekday()]))
    for a, b, c, k in linhas:
        html += '<tr><td style="color:#6b7383">%s</td><td align="right"><b>%s</b></td><td style="color:%s">%s</td></tr>' % (a, b, k, c)
    html += ('</table><p style="margin:20px 0 6px"><a href="%s" style="display:inline-block;background:#FFD100;color:#0F3C6E;'
             'text-decoration:none;font-weight:bold;padding:10px 18px;border-radius:6px">Abrir o painel</a></p>'
             '<p style="margin:0;color:#6b7383;font-size:12px">%s · acesso com o PIN de sempre. Painel já atualizado com os dados de %s.</p>'
             % (LINK, LINK, br(ref)))
    html += '</div>'
    txt = "Resultados de Vendas · Mercado Livre · marca Dellamed\nDados de %s (%s)\n\n" % (br(ref), SEM[ref.weekday()])
    txt += "\n".join(("%s: %s (%s)" % (a, b, c)) if a else ("   %s" % c) for a, b, c, _ in linhas)
    txt += "\n\nPainel atualizado: %s (acesso com o PIN de sempre)." % LINK
    html = "".join(c if ord(c) < 128 else "&#%d;" % ord(c) for c in html)
    return {"assunto": sub, "html": html, "texto": txt,
            "numeros": {"valor_dia": round(gd, 2), "pecas_dia": int(qd), "valor_mes_ate_ref": round(gm, 2), "vendedores_mes": vm, "anuncios_mes": am}}

# ---------------------------------------------------------------- comandos
def cmd_atualizar(index, master, ref, hoje, saida, pendentes='0', anterior=''):
    ref = datetime.date.fromisoformat(ref)
    rows = converter(master)
    if not rows: raise SystemExit('ERRO: o Excel do MASTER não tem linhas DELLAMED.')
    meses = sorted({r[1] for r in rows})
    if len(meses) != 1: raise SystemExit('ERRO: o Excel do MASTER traz mais de um mês: %s' % meses)
    if max(r[8] for r in rows) > ref.isoformat(): raise SystemExit('ERRO: há vendas depois de %s no Excel.' % ref)
    html = open(index, encoding='utf-8').read()
    keys = json.loads(_ler_const(html, 'EMBEDDED_KEYS'))
    html = gravar_meses(html, keys, {meses[0]: rows})
    open(index, 'w', encoding='utf-8').write(html)
    os.makedirs(saida, exist_ok=True)
    rm = resumo_mes(rows)
    json.dump(rm, open(os.path.join(saida, 'resumo_%s.json' % rm['mes']), 'w', encoding='utf-8'), separators=(',', ':'))
    dias = dict(ler_resumo_anterior(anterior)); dias.update(rm['dias'])
    em = montar_email(rows, {'dias': dias}, ref, int(pendentes or 0))
    json.dump(em, open(os.path.join(saida, 'email.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    print(json.dumps({'mes': meses[0], 'linhas': len(rows), 'valor': round(sum(r[5] for r in rows), 2),
                      'pecas': sum(r[6] for r in rows), 'ultima_data': max(r[8] for r in rows),
                      'assunto_resumo': em['assunto'], 'resumo_do_mes': 'resumo_%s.json' % rm['mes'],
                      'base_mes_anterior': bool(anterior and os.path.exists(anterior)), **em['numeros']}, ensure_ascii=False))

def cmd_conferir(index, pin):
    html = open(index, encoding='utf-8').read()
    keys = json.loads(_ler_const(html, 'EMBEDDED_KEYS'))
    priv = abrir_privada(keys, pin)
    for m, b in json.loads(_ler_const(html, 'EMBEDDED_MONTHS')).items():
        rows = destrancar(priv, b)
        print(m, len(rows), round(sum(r[5] for r in rows), 2))

if __name__ == '__main__':
    a = sys.argv[1:]
    if a and a[0] == 'chaves':
        json.dump(gerar_chaves(a[1]), open(a[2], 'w'))
    elif a and a[0] == 'atualizar':
        cmd_atualizar(*a[1:])
    elif a and a[0] == 'conferir':
        cmd_conferir(a[1], a[2])
    elif a and a[0] == 'montar':
        html = open(a[1], encoding='utf-8').read()
        keys = json.load(open(a[2])); linhas = json.load(open(a[3], encoding='utf-8'))
        por = collections.defaultdict(list)
        for r in linhas: por[r[1]].append(r)
        html = gravar_meses(html, keys, dict(por))
        open(a[1], 'w', encoding='utf-8').write(html)
    else:
        print(__doc__); sys.exit(1)
