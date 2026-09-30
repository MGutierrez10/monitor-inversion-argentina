"""Genera ied_cubo.json desde el Excel del BCRA "Información sobre inversiones directas en empresas residentes".
Uso: python3 make_ied.py <ruta_al_xlsx>
Usa 'Tabla de datos I' (posición pasiva bruta por país x sector x subsector) y
'Tabla de datos IV' (flujos transaccionales por país x sector x subsector, con componentes).
"""
import sys, json, os
import pandas as pd

f = sys.argv[1]
D = os.path.dirname(os.path.abspath(__file__))

SEC = {'Actividades administrativas y servicios de apoyo': 'Servicios administrativos', 'Agricultura, ganadería, caza, silvicultura y pesca': 'Agro y pesca', 'Auxiliares financieros': 'Auxiliares financieros', 'Comercio al por mayor y al por menor, reparación de vehículos automotores y motocicletas': 'Comercio', 'Confidencial': 'Confidencial', 'Construcción': 'Construcción', 'Enseñanza': 'Enseñanza', 'Explotación de minas y canteras': 'Minería y petróleo', 'Industria manufacturera': 'Industria manufacturera', 'Información y comunicaciones': 'Información y comunicaciones', 'Otros bienes o servicios prestados por personas jurídicas n.c.p.': 'Otros servicios n.c.p.', 'Otros intermediarios financieros': 'Otros intermediarios financieros', 'Salud humana y servicios sociales': 'Salud', 'Seguros': 'Seguros', 'Servicio de transporte y almacenamiento': 'Transporte y almacenamiento', 'Servicios artísticos, culturales, deportivos y de esparcimiento': 'Cultura y esparcimiento', 'Servicios de alojamiento y servicios de comida': 'Hotelería y gastronomía', 'Servicios de asociaciones y servicios personales': 'Asociaciones y serv. personales', 'Servicios inmobiliarios': 'Inmobiliario', 'Servicios profesionales, científicos y técnicos': 'Servicios profesionales', 'Sociedades captadoras de depósitos, excepto el banco central': 'Bancos', 'Suministro de agua, cloacas, gestión de residuos, recuperación de materiales y saneamiento público': 'Agua y saneamiento', 'Suministro de electricidad, gas, vapor y aire acondicionado': 'Electricidad y gas'}
PAIS = {'Estados Unidos De América': 'Estados Unidos', 'Reino Unido De Gran Bretaña E Irlanda Del Norte': 'Reino Unido', 'Republica De Corea (Corea Del Sur)': 'Corea del Sur', 'Islas Virgenes, Britanicas': 'Islas Vírgenes Británicas', 'República Bolivariana De Venezuela': 'Venezuela', 'Taiwán (Provincia De China)': 'Taiwán', 'Federación De Rusia': 'Rusia', 'Republica  Dominicana': 'República Dominicana', 'Nueva Zelandia': 'Nueva Zelanda', 'Sin Clasificar': 'Sin clasificar'}

def pais(x):
    x = ' '.join(str(x).split()) if str(x) not in PAIS else str(x)
    if x in PAIS: return PAIS[x]
    t = x.title()
    for a, b in ((' De ', ' de '), (' Y ', ' y '), (' Del ', ' del '), (' La ', ' la ')):
        t = t.replace(a, b)
    return t

def sub(x):
    x = ' '.join(str(x).split())
    return x[:1].upper() + x[1:]

t1 = pd.read_excel(f, sheet_name='Tabla de datos I', header=None, skiprows=7).iloc[:, :7]
t1.columns = ['per', 'pais', 'sec', 'sub', 'pos', 'cap', 'deu']
t4 = pd.read_excel(f, sheet_name='Tabla de datos IV', header=None, skiprows=7).iloc[:, :10]
t4.columns = ['per', 'pais', 'sec', 'sub', 'tot', 'apo', 'fya', 'div', 'ren', 'deuf']
for t in (t1, t4):
    t.dropna(subset=['per'], inplace=True)
    t['q'] = pd.to_datetime(t['per']).dt.to_period('Q').astype(str)
    t['pais'] = t['pais'].map(pais)
    t['sec'] = t['sec'].map(lambda s: SEC.get(' '.join(str(s).split()), ' '.join(str(s).split())))
    t['sub'] = t['sub'].map(sub)
    t.drop(columns='per', inplace=True)
k = ['q', 'pais', 'sec', 'sub']
t1 = t1.groupby(k, as_index=False).sum(numeric_only=True)
t4 = t4.groupby(k, as_index=False).sum(numeric_only=True)
m = t1.merge(t4, on=k, how='outer').fillna(0.0)
vals = ['pos', 'cap', 'deu', 'tot', 'apo', 'fya', 'div', 'ren', 'deuf']
m = m[(m[vals].abs() >= 0.05).any(axis=1)]

Q = sorted(m['q'].unique())
P = m.groupby('pais')['pos'].sum().sort_values(ascending=False).index.tolist()
S = m.groupby('sec')['pos'].sum().sort_values(ascending=False).index.tolist()
U = sorted(m[['sec', 'sub']].drop_duplicates().itertuples(index=False), key=lambda r: (S.index(r[0]), r[1]))
Ui = {(s, u): i for i, (s, u) in enumerate(U)}
qi = {q: i for i, q in enumerate(Q)}; pi = {p: i for i, p in enumerate(P)}

def r1(x):
    x = round(float(x), 1)
    return int(x) if x == int(x) else x

rows = [[qi[r.q], pi[r.pais], Ui[(r.sec, r.sub)]] + [r1(getattr(r, v)) for v in vals] for r in m.itertuples(index=False)]

# Tablas de una sola apertura (II y V): son las que usa el BCRA en sus cuadros por sector o por país.
# Difieren de la tabla cruzada porque la confidencialidad (<3 empresas) se aplica por celda.
t2 = pd.read_excel(f, sheet_name='Tabla de datos II', header=None, skiprows=7)
t5 = pd.read_excel(f, sheet_name='Tabla de datos V', header=None, skiprows=7)
def prep(df, cols, key):
    df = df.iloc[:, cols].copy(); df.columns = ['per'] + key + (['pos', 'cap', 'deu'] if len(cols) - len(key) == 4 else ['tot', 'apo', 'fya', 'div', 'ren', 'deuf'])
    df = df.dropna(subset=['per']); df['q'] = pd.to_datetime(df['per']).dt.to_period('Q').astype(str); return df.drop(columns='per')
s2 = prep(t2, [0, 1, 2, 3, 4, 5], ['sec', 'sub']); p2 = prep(t2, [8, 9, 10, 11, 12], ['pais'])
s5 = prep(t5, [1, 2, 3, 4, 5, 6, 7, 8, 9], ['sec', 'sub']); p5 = prep(t5, [12, 13, 14, 15, 16, 17, 18, 19], ['pais'])
for d in (s2, s5):
    d['sec'] = d['sec'].map(lambda x: SEC.get(' '.join(str(x).split()), ' '.join(str(x).split()))); d['sub'] = d['sub'].map(sub)
for d in (p2, p5): d['pais'] = d['pais'].map(pais)
ss = s2.groupby(['q', 'sec', 'sub'], as_index=False).sum(numeric_only=True).merge(s5.groupby(['q', 'sec', 'sub'], as_index=False).sum(numeric_only=True), on=['q', 'sec', 'sub'], how='outer').fillna(0.0)
pp = p2.groupby(['q', 'pais'], as_index=False).sum(numeric_only=True).merge(p5.groupby(['q', 'pais'], as_index=False).sum(numeric_only=True), on=['q', 'pais'], how='outer').fillna(0.0)
for s_, u_ in ss[['sec', 'sub']].drop_duplicates().itertuples(index=False):
    if s_ not in S: S.append(s_)
    if (s_, u_) not in Ui: Ui[(s_, u_)] = len(U); U.append((s_, u_))
for p_ in pp['pais'].unique():
    if p_ not in pi: pi[p_] = len(P); P.append(p_)
ss = ss[(ss[vals].abs() >= 0.05).any(axis=1)]; pp = pp[(pp[vals].abs() >= 0.05).any(axis=1)]
rows_sec = [[qi[r.q], -1, Ui[(r.sec, r.sub)]] + [r1(getattr(r, v)) for v in vals] for r in ss.itertuples(index=False)]
rows_pais = [[qi[r.q], pi[r.pais], -1] + [r1(getattr(r, v)) for v in vals] for r in pp.itertuples(index=False)]
print('control sector-only 2026Q1:', ss[ss.q == '2026Q1'].groupby('sec')[['pos', 'tot']].sum().round(1).loc[['Minería y petróleo', 'Industria manufacturera']].to_dict())
out = {'fuente': 'BCRA, Información sobre inversiones directas en empresas residentes (Tablas de datos I y IV)',
       'trimestres': Q, 'paises': P, 'sectores': S, 'subsectores': [[S.index(s), u] for s, u in U],
       'campos': ['q', 'pais', 'sub'] + vals, 'filas': rows, 'filas_sec': rows_sec, 'filas_pais': rows_pais}
json.dump(out, open(os.path.join(D, 'ied_cubo.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
chk = m.groupby('q')[['pos', 'tot']].sum().tail(2).round(1)
print('ok', len(rows), 'filas', Q[0], '-', Q[-1], '| control últimos trimestres:\n', chk)
