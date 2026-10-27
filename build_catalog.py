#!/usr/bin/env python3
"""Transforma el Excel «Fórmula y lógica - Macro PPTO» en el catálogo compacto que embebe el prototipo.
Uso:  python3 build_catalog.py <archivo.xlsx> <salida.b64>
No modifica el Excel. Hoja fuente: «TABLA DATOS GENERALES» (Codigo, Descripcion, Tarifa; clave XX = Codigo+Sede+Financiador+Producto+Grupo).
Salida: gzip( u32 longitudJSON | JSON(meta) | por cada código: u16 n, n × (u16 contexto, u32 centimos) ) codificado en base64."""
import sys,json,gzip,struct,base64,collections,openpyxl
sys.path.insert(0,__import__('os').path.dirname(__file__))
from split_ctx import split_suffix
src,out=sys.argv[1],sys.argv[2]
wb=openpyxl.load_workbook(src,data_only=True,read_only=True)
ws=wb['TABLA DATOS GENERALES']
rows=[r for r in ws.iter_rows(min_row=2,values_only=True) if r[0] and r[1]]
desc={};tar=collections.defaultdict(list);ctxs={};stats=collections.Counter()
for r in rows:
    key,code,d,t=str(r[0]),str(r[1]).strip(),(r[2] or '').strip(),r[3]
    suf=key[len(code):]
    if suf not in ctxs:ctxs[suf]=len(ctxs)
    desc.setdefault(code,d)
    if t is None:stats['sin_monto']+=1;continue
    c=round(float(t)*100)
    if (ctxs[suf],c) in tar[code]:stats['duplicado_identico']+=1;continue
    if any(x[0]==ctxs[suf] for x in tar[code]):stats['duplicado_monto_distinto']+=1
    tar[code].append((ctxs[suf],c))
codes=sorted(desc)
ctx_list=[None]*len(ctxs)
for s,i in ctxs.items():
    f,p,g=split_suffix(s);ctx_list[i]=[f,p,g.replace('Grupo Telefonica / ','').replace('NO ESPECIFICO / ','').strip(' /') if g else '']
meta={'sede':'Clínica Delgado','ctx':ctx_list,'codes':[[c,desc[c]] for c in codes],'filas':len(rows),'stats':dict(stats)}
mj=json.dumps(meta,ensure_ascii=False,separators=(',',':')).encode()
b=bytearray(struct.pack('<I',len(mj)))+mj
for c in codes:
    t=tar.get(c,[]);b+=struct.pack('<H',len(t))
    for ci,cs in t:b+=struct.pack('<HI',ci,cs)
z=gzip.compress(bytes(b),9)
open(out,'w').write(base64.b64encode(z).decode())
print('filas',len(rows),'codigos',len(codes),'contextos',len(ctxs),'bytes bin',len(b),'gz',len(z),'b64',len(z)*4//3,dict(stats))
