import csv, re, json, unicodedata, sys, openpyxl
from collections import defaultdict
CSV, XLSX, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
rows = list(csv.reader(open(CSV, encoding='latin-1'), delimiter=';'))

def strip_acc(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
def key(s):
    s = strip_acc(s).upper().replace('.', ' ').replace('-', ' ')
    return re.sub(r'\s+', ' ', s).strip()

# ---------- estatisticas da planilha ----------
mcs = []
for r in rows[1:]:
    if not r[1].strip(): continue
    num = lambda x: float(x.replace(',', '.')) if x.strip() else 0
    mcs.append(dict(uf=r[0].strip(), raw_uf=r[0], nome=r[1].strip(), raw_nome=r[1],
        sheet=dict(edicoes=num(r[2]), titulos=num(r[3]), vices=num(r[4]), batalhas=num(r[5]),
                   vitorias=num(r[6]), media=num(r[7]), rounds=num(r[8]), rounds_v=num(r[9]),
                   media_r=num(r[10]), twola_d=num(r[11]), twola_t=num(r[12]))))
namecount = defaultdict(int)
for m in mcs: namecount[key(m['nome'])] += 1
for m in mcs:
    k = key(m['nome'])
    m['id'] = k if namecount[k] == 1 else f"{k}|{m['uf']}"
byid = {m['id']: m for m in mcs}
ALIAS = {'TAGRELA': 'TAGARELA', 'KOEL': 'KOELL', 'MARTZIN': 'MARTIZIN', 'VINCIUS ZN': 'VINICIUS ZN'}
inconsist = []
def add_inc(cat, sev, msg, ano=None):
    inconsist.append(dict(cat=cat, sev=sev, msg=msg, ano=ano))

alias_seen = set()
def resolve(raw, ano):
    s = raw.strip()
    m = re.match(r'^(.*?)\s*\(?([A-Z]{2})\)\s*$', s)
    uf_tag = None
    if m: s, uf_tag = m.group(1), m.group(2)
    if re.search(r'\([A-Z]{2}$', raw.strip()) or (raw.strip().endswith(')') and '(' not in raw):
        add_inc('Digitação', 'baixa', f'Estado mal formatado no chaveamento: "{raw.strip()}"', ano)
    k = key(s)
    if k in ALIAS:
        if (k, ano) not in alias_seen:
            add_inc('Grafia', 'baixa', f'Nome "{s.strip()}" no chaveamento de {ano} não bate com a planilha de estatísticas — tratado como "{byid[ALIAS[k]]["nome"]}"', ano)
            alias_seen.add((k, ano))
        k = ALIAS[k]
    cid = k
    if cid not in byid:
        if uf_tag and f'{k}|{uf_tag}' in byid: cid = f'{k}|{uf_tag}'
        else:
            add_inc('MC não encontrado', 'alta', f'"{raw.strip()}" ({ano}) não existe na planilha de estatísticas', ano)
            byid[cid] = dict(id=cid, nome=s.strip(), uf='??', sheet=None)
            mcs.append(byid[cid])
    if uf_tag and byid[cid]['uf'] != uf_tag:
        add_inc('Estado divergente', 'alta', f'{byid[cid]["nome"]} aparece como ({uf_tag}) no chaveamento de {ano}, mas é {byid[cid]["uf"]} na planilha', ano)
    return cid

# ---------- chaveamentos ----------
battles = []; champs_line = {}
PH = {'SEMI FINAL': 'Semifinal', 'FINAL': 'Final', '2ª FASE': '2ª fase', 'QUARTAS DE FINAL': 'Quartas de final'}
for c0 in (14, 20, 26):
    ano = None; fase = None; order = 0
    for ri in range(1, len(rows)):
        r = rows[ri] + [''] * 50
        a = r[c0].strip()
        if re.fullmatch(r'20\d\d', a):
            ano = int(a); fase = '1ª fase'; continue
        if ano is None: continue
        if a.upper() in PH or a == '2Âª FASE' or 'FASE' in a.upper():
            fase = PH.get(a.upper(), '2ª fase'); continue
        if not a and r[c0+1].strip() and not r[c0+2].strip():
            champs_line[ano] = r[c0+1].strip(); continue
        if a and r[c0+2].strip().upper() == 'X':
            order += 1
            sa, sb, b = int(r[c0+1]), int(r[c0+3]), r[c0+4]
            if '/' in b:
                others = [x for x in b.split('/')]
                battles.append(dict(ano=ano, fase=fase, ordem=order, tipo='triplo', raw=f'{a} {sa}x{sb} {b}',
                    a=resolve(a, ano), others=[resolve(x, ano) for x in others]))
            else:
                battles.append(dict(ano=ano, fase=fase, ordem=order, tipo='normal', a=resolve(a, ano), sa=sa, sb=sb, b=resolve(b, ano)))

# triplo 2017: Menor 1x0 Miliano/Blackout + Miliano 2x0 Blackout => uma batalha
final_b = []
for bt in battles:
    if bt['tipo'] == 'triplo':
        nxt = next(x for x in battles if x['ano'] == bt['ano'] and x['ordem'] == bt['ordem'] + 1)
        nxt['skip'] = True
        first, (second, third) = bt['a'], bt['others']
        bt['st'] = {first: [1, 1, 1, 0, 0], nxt['a']: [1, 3, 2, 1, 0], nxt['b']: [0, 3, 0, 0, 1]}; bt['rounds'] = 3
        bt.update(participantes=[first, second, third], classificados=[first, nxt['a']],
                  eliminado=nxt['b'], placar=f"{bt['raw']} · {byid[nxt['a']]['nome']} {nxt['sa']}x{nxt['sb']} {byid[nxt['b']]['nome']}",
                  pares=[[first, second, 1, 0], [first, third, 1, 0], [nxt['a'], nxt['b'], nxt['sa'], nxt['sb']]])
        bt['_n'] = nxt
# double three de 2016 (não está na planilha de chaveamento): Vieira caiu no 1º round; TH e Nicolas Walter seguiram
PRELIM = [dict(ano=2016, fase='Fase preliminar', ordem=0, tipo='triplo', participantes=['TH', 'NICOLAS WALTER', 'VIEIRA'],
               classificados=['TH', 'NICOLAS WALTER'], eliminado='VIEIRA', rounds=1,
               placar='Double three: Vieira eliminado no 1º round; TH e Nicolas Walter classificados',
               pares=[['TH', 'VIEIRA', 1, 0], ['NICOLAS WALTER', 'VIEIRA', 1, 0]],
               st={'TH': [1, 1, 1, 0, 0], 'NICOLAS WALTER': [1, 1, 1, 0, 0], 'VIEIRA': [0, 1, 0, 0, 0]})]
for bt in battles:
    if bt['ano'] == 2016 and PRELIM:
        final_b.extend(PRELIM); PRELIM = []
    if not bt.get('skip'): final_b.append(bt)
battles = final_b
for i, bt in enumerate(battles): bt['id'] = i

for bt in battles:
    if bt['tipo'] == 'normal':
        if bt['sa'] < bt['sb']:
            add_inc('Placar', 'média', f"{bt['ano']} {bt['fase']}: placar invertido ({byid[bt['a']]['nome']} {bt['sa']}x{bt['sb']} {byid[bt['b']]['nome']}) — vencedor deveria vir à esquerda", bt['ano'])
            bt['a'], bt['b'], bt['sa'], bt['sb'] = bt['b'], bt['a'], bt['sb'], bt['sa']
        if bt['sa'] != 2 or bt['sb'] > 1:
            add_inc('Placar', 'média', f"{bt['ano']} {bt['fase']}: placar incomum {byid[bt['a']]['nome']} {bt['sa']}x{bt['sb']} {byid[bt['b']]['nome']}", bt['ano'])

years = sorted({b['ano'] for b in battles})
for y in years:
    first = [b for b in battles if b['ano'] == y and b['fase'] == '1ª fase']
    if len(first) == 4:
        for b in first: b['fase'] = 'Quartas de final'
# progressão do chaveamento
for y in years:
    bs = [b for b in battles if b['ano'] == y]
    fases = []
    for b in bs:
        if b['fase'] not in fases: fases.append(b['fase'])
    for f1, f2 in zip(fases, fases[1:]):
        if f1 == 'Fase preliminar': continue
        win = set(); 
        for b in bs:
            if b['fase'] == f1:
                if b['tipo'] == 'triplo': win |= set(b['classificados'])
                else: win.add(b['a'])
        part = set()
        for b in bs:
            if b['fase'] == f2: part |= {b['a'], b['b']}
        for x in part - win:
            add_inc('Chaveamento', 'alta', f'{y}: {byid[x]["nome"]} aparece em {f2} sem ter vencido em {f1}', y)
        for x in win - part:
            add_inc('Chaveamento', 'alta', f'{y}: {byid[x]["nome"]} venceu em {f1} mas não aparece em {f2}', y)
    fin = [b for b in bs if b['fase'] == 'Final']
    cl = champs_line.get(y)
    if fin and cl:
        cid = resolve(cl, y)
        if cid != fin[0]['a']:
            add_inc('Campeão', 'alta', f'{y}: campeão escrito abaixo da final ({cl}) difere do vencedor da final ({byid[fin[0]["a"]]["nome"]})', y)
    # repetidos na mesma fase
    for f in fases:
        seen = defaultdict(int)
        for b in bs:
            if b['fase'] == f:
                for p in (b.get('participantes') or [b['a'], b['b']]): seen[p] += 1
        for p, n in seen.items():
            if n > 1: add_inc('Chaveamento', 'alta', f'{y} {f}: {byid[p]["nome"]} aparece {n} vezes', y)

# ---------- estaduais ----------
wb = openpyxl.load_workbook(XLSX, data_only=True); ws = wb.active
allrows = list(ws.iter_rows(values_only=True))
hi = next(i for i, r in enumerate(allrows) if 2012 in r or 2012.0 in r)
hdr = allrows[hi]
ycols = [(i, (2021 if str(v) == '2020/2021' else int(v)), str(v if str(v) == '2020/2021' else int(v))) for i, v in enumerate(hdr) if v]
UFN = {}
estaduais = []
CH_ALIAS = {'BIRO BRIO': 'BIRO BIRO', 'KADOSH': 'KADOSHI', 'ALESON': 'ALESSON'}
for row in allrows[hi+1:]:
    if not row[0] or '(' not in str(row[0]): continue
    m = re.match(r'(.*)\((\w\w)\)', row[0]); uf = m.group(2); UFN[uf] = m.group(1).strip()
    for i, y, lab in ycols:
        v = row[i]
        if v is None or str(v).strip() in ('-', ''): continue
        nm = str(v).strip()
        if isinstance(v, float): nm = str(int(v))
        k = key(nm); k2 = CH_ALIAS.get(k, k)
        mid = None
        if k2 in byid: mid = k2
        elif f'{k2}|{uf}' in byid: mid = f'{k2}|{uf}'
        estaduais.append(dict(uf=uf, ano=y, rot=lab, nome=nm, mc=mid))
key_uf = {}
HOMONIMOS = {('SE', 'NG'), ('DF', 'JEFFE'), ('SC', 'VITU')}  # MA Jeffe ≠ DF Jeffe; PE Vitu ≠ SC Vitu; MG NG ≠ SE NG  # confirmados: MCs diferentes

# participantes por ano
part = defaultdict(set)
for b in battles:
    for p in (b.get('participantes') or [b['a'], b['b']]): part[b['ano']].add(p)

for e in estaduais:
    mid = e['mc']
    if mid and byid[mid]['uf'] != e['uf']:
        if e['ano'] in years and mid in part[e['ano']]:
            add_inc('Estado divergente', 'alta', f"{e['nome']} é campeão estadual de {e['uf']} em {e['rot']}, mas está como {byid[mid]['uf']} na planilha", e['ano'])
        else:
            if (e['uf'], key(e['nome'])) not in HOMONIMOS:
                add_inc('Homônimo?', 'baixa', f"Campeão estadual {e['uf']} {e['rot']} \"{e['nome']}\" tem o mesmo nome de {byid[mid]['nome']} ({byid[mid]['uf']}) — tratado como outro MC", e['ano'])
            e['mc'] = None
    if e['ano'] in years:
        if not e['mc'] or e['mc'] not in part[e['ano']]:
            e['no_nacional'] = True
    elif e['ano'] not in years:
        e['sem_dados'] = True
# grafias diferentes estaduais x estatísticas
for e in estaduais:
    if e['mc']:
        if key(e['nome']) != e['mc'].split('|')[0]:
            add_inc('Nome alternativo', 'info', f"Estaduais {e['uf']} {e['rot']}: \"{e['nome']}\" ligado a \"{byid[e['mc']]['nome']}\" das estatísticas (nome completo/variação)", e['ano'])

est_by = {(e['uf'], e['ano']): e for e in estaduais}
for y in years:
    for p in part[y]:
        uf = byid[p]['uf']
        e = est_by.get((uf, y))
        if not e or e['mc'] != p:
            pass  # vaga extra — informativo
vagas = {}
for y in years:
    vagas[y] = {}
    for p in part[y]:
        e = est_by.get((byid[p]['uf'], y))
        vagas[y][p] = 'estadual' if e and e['mc'] == p else ('repescagem' if y >= 2021 else 'outra')

# ---------- stats calculadas ----------
calc = {m['id']: dict(edicoes=set(), titulos=0, vices=0, batalhas=0, vitorias=0, rounds=0, rounds_v=0, twola_d=0, twola_t=0) for m in mcs}
for b in battles:
    if b['tipo'] == 'triplo':
        for p, (w, r, rv, td, tt) in b['st'].items():
            c = calc[p]; c['edicoes'].add(b['ano']); c['batalhas'] += 1; c['vitorias'] += w; c['rounds'] += r; c['rounds_v'] += rv; c['twola_d'] += td; c['twola_t'] += tt
        continue
    for p, me, op in ((b['a'], b['sa'], b['sb']), (b['b'], b['sb'], b['sa'])):
        c = calc[p]; c['edicoes'].add(b['ano']); c['batalhas'] += 1; c['vitorias'] += me > op
        c['rounds'] += me + op; c['rounds_v'] += me
        if me == 2 and op == 0: c['twola_d'] += 1
        if op == 2 and me == 0: c['twola_t'] += 1
    if b['fase'] == 'Final':
        calc[b['a']]['titulos'] += 1; calc[b['b']]['vices'] += 1
LAB = dict(edicoes='Edições', titulos='Títulos', vices='Vices', batalhas='Batalhas', vitorias='Vitórias', rounds='Rounds', rounds_v='Rounds vencidos', twola_d='Deu twolala', twola_t='Tomou twolala')
for m in mcs:
    c = calc[m['id']]; c['edicoes'] = len(c['edicoes'])
    m['calc'] = {k: c[k] for k in LAB}
    s = m.get('sheet')
    if s:
        diffs = [f"{LAB[k]}: planilha {int(s[k])} × chaveamento {c[k]}" for k in LAB if int(s[k]) != c[k]]
        if diffs:
            add_inc('Estatística', 'média', f"{m['nome']} ({m['uf']}): " + '; '.join(diffs))
            m['diffs'] = diffs
        if s['batalhas'] and abs(s['media'] - 100 * s['vitorias'] / s['batalhas']) > 0.1:
            add_inc('Estatística', 'baixa', f"{m['nome']}: média de vitórias da planilha ({s['media']:.1f}%) não confere com vitórias/batalhas")
        if s['rounds'] and abs(s['media_r'] - 100 * s['rounds_v'] / s['rounds']) > 0.1:
            add_inc('Estatística', 'baixa', f"{m['nome']}: média de rounds da planilha ({s['media_r']:.1f}%) não confere")
    if m.get('raw_uf') and m['raw_uf'] != m['raw_uf'].strip():
        add_inc('Digitação', 'baixa', f"{m['nome']}: estado com espaço extra (\"{m['raw_uf']}\") na planilha")
    if m.get('raw_nome') and m['raw_nome'] != m['raw_nome'].strip():
        add_inc('Digitação', 'baixa', f"\"{m['raw_nome']}\": nome com espaço extra na planilha")
    if m.get('sheet') and c['batalhas'] == 0:
        add_inc('MC sem batalha', 'alta', f"{m['nome']} ({m['uf']}) está na planilha com {int(m['sheet']['batalhas'])} batalha(s), mas não aparece em nenhum chaveamento")
    m.pop('raw_uf', None); m.pop('raw_nome', None)

for b in battles:
    b.pop('_n', None); b.pop('raw', None)
champions = {}
for y in years:
    f = next(b for b in battles if b['ano'] == y and b['fase'] == 'Final')
    champions[y] = dict(campeao=f['a'], vice=f['b'], placar=f"{f['sa']}x{f['sb']}")
sev_order = {'alta': 0, 'média': 1, 'baixa': 2, 'info': 3}
inconsist.sort(key=lambda x: (sev_order[x['sev']], x['cat'], x['ano'] or 0))
data = dict(ufs=UFN, mcs=[{k: v for k, v in m.items()} for m in mcs], battles=battles, years=years,
            champions=champions, vagas=vagas, estaduais=estaduais, est_years=[[y, l] for _, y, l in ycols], inconsist=inconsist)
open(OUT, 'w', encoding='utf-8').write('window.DUELO = ' + json.dumps(data, ensure_ascii=False) + ';\n')
print(len(mcs), 'mcs', len(battles), 'batalhas', years)
for i in inconsist: print(i['sev'], '|', i['cat'], '|', i['msg'])
