#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AltaSala · comprobación profunda de currículums contra datos públicos.

Uso:
  APIFY_TOKEN=xxx python3 verify_cv.py carpeta_con_cvs/ --ciudad Barcelona --out informe.json

Qué hace con cada PDF o DOCX:
  1. Extrae el texto (pdftotext si está instalado; si no, pypdf; para .docx usa python-docx).
  2. Saca nombre, teléfono, correo, puestos con fechas, escuelas e idiomas.
  3. Comprueba cada casa en OpenStreetMap (Nominatim, 1 petición por segundo, con User-Agent).
  4. Busca a la persona en LinkedIn con el actor harvestapi/linkedin-profile-search de Apify
     (nombre + ciudad) y compara puesto y casa con lo que pone el currículum.
  5. Comprueba fechas (solapes, huecos, años declarados contra años sumados).
  6. Escribe un informe JSON con cada afirmación y su fuente, y un resumen en pantalla.

No inventa nada: lo que no se encuentra sale como "no encontrado". La aceptación del puesto
solo existe tras la llamada; este script no la simula.
"""
import os,sys,re,json,time,argparse,subprocess,unicodedata,urllib.request,urllib.parse

OSM_OK=re.compile(r'restaurant|bar|cafe|pub|fast_food|nightclub|biergarten|food_court|hotel|college|school|university')
ROLES=[('Jefe de cocina',r'jefe de cocina|chef ejecutivo|head chef'),('Segundo de cocina',r'segundo de cocina|sous'),('Jefe de partida',r'jefe de partida|cocinero|chef de partie|cook|sushiman'),('Jefe de sala',r'jefe de sala|jefa de sala|maitre|maître|director de sala|encargad'),('Camarero',r'camarer|jefe de rango|jefa de rango|waiter|runner'),('Sumiller',r'sumiller|sommelier'),('Barra',r'barra|bartender|barman|barista|coctel')]
SCHOOLS=json.load(open(os.path.join(os.path.dirname(__file__),'schools.json'),encoding='utf-8')) if os.path.exists(os.path.join(os.path.dirname(__file__),'schools.json')) else {'schools':[],'titles':[]}

def deacc(s): return ''.join(c for c in unicodedata.normalize('NFD',s or '') if unicodedata.category(c)!='Mn')
def nrm(s): return re.sub(r'[^a-z0-9]+','',re.sub(r'\b(restaurante?|restaurant|grupo|group|hotel|bar|the|el|la|los|las|de|del|by|sl|madrid|barcelona)\b',' ',deacc(s).lower()))

def extract_text(path):
    if path.lower().endswith('.pdf'):
        try: return subprocess.run(['pdftotext','-layout',path,'-'],capture_output=True,text=True,check=True).stdout
        except Exception:
            from pypdf import PdfReader
            return '\n'.join((p.extract_text() or '') for p in PdfReader(path).pages)
    if path.lower().endswith('.docx'):
        import docx
        return '\n'.join(p.text for p in docx.Document(path).paragraphs)
    return open(path,encoding='utf-8',errors='ignore').read()

def parse(text):
    lines=[re.sub(r'\s+',' ',l).strip() for l in text.splitlines()];lines=[l for l in lines if l]
    cv={'name':'','phone':'','email':'','exps':[],'years':None,'schools':[],'langs':[]}
    m=re.search(r'[\w.+-]+@[\w-]+\.[\w.]+',text);cv['email']=m.group(0) if m else ''
    m=re.search(r'(?:\+34\s?)?[67]\d{2}[\s.]?\d{3}[\s.]?\d{3}\b',text);cv['phone']=re.sub(r'[\s.]','',m.group(0)) if m else ''
    for l in lines[:4]:
        w=l.split(' ')
        if 2<=len(w)<=5 and not re.search(r'\d|@|·',l) and all(re.match(r"^[A-ZÁÉÍÓÚÑ][\wáéíóúñü'-]*$|^(de|del|la|el|y)$",x) for x in w): cv['name']=l;break
    if not cv['name']: cv['name']=lines[0].split('·')[0].strip()
    m=re.search(r'(\d{1,2})\s*años? de experiencia',text,re.I);cv['years']=int(m.group(1)) if m else None
    rx=re.compile(r'(\d{4})\s*[-–—a]\s*(\d{4}|actualidad|presente|hoy|act\.?)',re.I)
    for l in lines:
        m=rx.search(l)
        if not m: continue
        if re.search(r'escuela|escola|cordon|hofmann|cett|universi|instituto|grado|t[eé]cnico|diploma|m[aá]ster|wset|manipulador',l,re.I) and not any(re.search(p,l,re.I) for _,p in ROLES): continue
        a=int(m.group(1));b=int(m.group(2)) if m.group(2).isdigit() else time.localtime().tm_year
        segs=[s.strip() for s in re.split(r'·|\||\s[-–—]\s',l.replace(m.group(0),'')) if s.strip()]
        role=next((s for s in segs if any(re.search(p,s,re.I) for _,p in ROLES) and len(s)<40),'')
        venue=next((re.sub(r',\s*(Madrid|Barcelona|Valencia|Sevilla).*$','',s).strip() for s in segs if s!=role and not re.match(r'^(madrid|barcelona|\d)',s,re.I)),'')
        cv['exps'].append({'role':role,'venue':venue,'from':a,'to':b,'current':not m.group(2).isdigit()})
    low=deacc(text).lower()
    for s in SCHOOLS.get('schools',[])+SCHOOLS.get('titles',[]):
        if deacc(s).lower() in low: cv['schools'].append(s)
    for m in re.finditer(r'\b(ingl[eé]s|english|franc[eé]s|catal[aá]n|alem[aá]n|italiano|[aá]rabe|portugu[eé]s|chino)\b\s*(A1|A2|B1|B2|C1|C2|nativ[oa]|biling[üu]e|alto|medio|b[aá]sico)?',text,re.I):
        cv['langs'].append((m.group(1).capitalize()+' '+(m.group(2) or '')).strip())
    return cv

def osm(name,city):
    q=urllib.parse.urlencode({'q':f'{name}, {city}','format':'jsonv2','limit':3,'accept-language':'es'})
    req=urllib.request.Request('https://nominatim.openstreetmap.org/search?'+q,headers={'User-Agent':'AltaSala verify_cv (hugo@vystral.com)'})
    try: r=json.load(urllib.request.urlopen(req,timeout=20))
    except Exception as e: return {'found':None,'error':str(e)}
    time.sleep(1.1)
    h=next((x for x in r if OSM_OK.search(x.get('type',''))),None)
    return {'found':bool(h),'name':h.get('name') if h else None,'type':h.get('type') if h else None,'addr':', '.join(h.get('display_name','').split(',')[:2]) if h else None}

def linkedin(name,city,token):
    if not token: return {'found':None,'error':'sin APIFY_TOKEN'}
    url=f'https://api.apify.com/v2/acts/harvestapi~linkedin-profile-search/run-sync-get-dataset-items?token={token}&timeout=120'
    body=json.dumps({'searchQuery':name,'locations':[city],'maxItems':3,'profileScraperMode':'Short'}).encode()
    req=urllib.request.Request(url,data=body,headers={'Content-Type':'application/json'})
    try: items=json.load(urllib.request.urlopen(req,timeout=180))
    except Exception as e: return {'found':None,'error':str(e)}
    hits=[p for p in items if nrm((p.get('firstName','')+' '+p.get('lastName','')))==nrm(name)]
    if not hits: return {'found':False}
    p=hits[0];cp=(p.get('currentPositions') or [{}])[0]
    return {'found':True,'url':p.get('linkedinUrl'),'title':cp.get('title'),'company':cp.get('companyName'),'tenure':cp.get('tenureAtPosition')}

def verify(cv,city,token):
    out=[];add=lambda st,t,d,src:out.append({'estado':st,'dato':t,'detalle':d,'fuente':src})
    for e in cv['exps']:
        if not e['venue']: continue
        o=osm(e['venue'],city)
        if o.get('found'): add('ok',e['venue'],f"Existe: {o['name']}, {o['addr']}",'OpenStreetMap')
        elif o.get('found') is False: add('no',e['venue'],'No la encontramos en el mapa. Preguntar en la llamada','OpenStreetMap')
        else: add('pendiente',e['venue'],o.get('error','sin respuesta'),'OpenStreetMap')
    ex=[e for e in cv['exps'] if e['from']];total=sum(max(0,e['to']-e['from']) for e in ex)
    for i in range(len(ex)):
        for j in range(i+1,len(ex)):
            a,b=ex[i],ex[j];ov=min(a['to'],b['to'])-max(a['from'],b['from'])
            if ov>=1 and not (a['current'] and b['current']): add('aviso','Fechas',f"{a['venue']} y {b['venue']} se pisan {ov} año(s)",'fechas del currículum')
    if cv['years'] is not None:
        add('aviso' if cv['years']>total+1 else 'ok','Años declarados',f"Dice {cv['years']}, las fechas suman {total}",'fechas del currículum')
    if cv['schools']: add('ok','Formación',', '.join(cv['schools'])+' (reconocida)','catálogo de escuelas y títulos')
    else: add('pendiente','Formación','Sin formación reglada en el currículum','currículum')
    add('ok' if cv['phone'] else 'no','Teléfono',cv['phone'] or 'Sin teléfono: no se puede llamar','currículum')
    li=linkedin(cv['name'],city,token)
    if li.get('found'):
        same=any(nrm(e['venue'])==nrm(li.get('company') or '') for e in cv['exps'])
        add('ok' if same else 'aviso','LinkedIn',f"{li.get('title')} en {li.get('company')} · {li.get('url')}"+('' if same else ' · no coincide con el currículum'),'LinkedIn')
    elif li.get('found') is False: add('pendiente','LinkedIn','Sin perfil. Normal en sala y cocina. Referencia por teléfono a su última casa','LinkedIn')
    else: add('pendiente','LinkedIn',li.get('error',''),'LinkedIn')
    if cv['langs']: add('pendiente','Idiomas',', '.join(cv['langs'])+' · se comprueban hablando en la llamada','la llamada')
    return out

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('carpeta');ap.add_argument('--ciudad',default='Madrid');ap.add_argument('--out',default='informe_cv.json');a=ap.parse_args()
    token=os.environ.get('APIFY_TOKEN','');rep=[]
    for f in sorted(os.listdir(a.carpeta)):
        if not re.search(r'\.(pdf|docx|txt)$',f,re.I): continue
        cv=parse(extract_text(os.path.join(a.carpeta,f)));checks=verify(cv,a.ciudad,token)
        ok=sum(c['estado']=='ok' for c in checks);tot=sum(c['estado'] in('ok','no','aviso') for c in checks)
        rep.append({'archivo':f,'nombre':cv['name'],'telefono':cv['phone'],'puestos':cv['exps'],'comprobados':f'{ok} de {tot}','comprobaciones':checks})
        print(f"{cv['name'] or f}: {ok} de {tot} comprobados · "+' · '.join(f"{c['dato']} {'✓' if c['estado']=='ok' else '✗' if c['estado']=='no' else '!' if c['estado']=='aviso' else '…'}" for c in checks))
    json.dump(rep,open(a.out,'w',encoding='utf-8'),ensure_ascii=False,indent=1);print('informe:',a.out)
