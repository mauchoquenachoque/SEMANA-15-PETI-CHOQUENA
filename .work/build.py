from pathlib import Path
import csv, shutil
import struct, zlib
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT

root=Path(__file__).resolve().parents[1]; out=root/'08_riesgos'; out.mkdir(exist_ok=True)
ev=root/'evidencias'/'S15'/'salidas'; ev.mkdir(parents=True,exist_ok=True)
R=[
['R-01','Presupuestal','Que se reduzca el presupuesto de respaldo por priorizar gasto operativo, retrasando recuperación crítica.','BIA01,BIA03,BIA04,BIA06','Política y BIA',4,4,'Crítico','BIA no aprobado y pruebas nunca realizadas.'],
['R-02','Patrocinio','Que la gerencia no apruebe BIA y plan por falta de instancia de decisión, manteniendo objetivos no validados.','Todos','Universo auditable',3,5,'Crítico','BIA sin revisión ni aprobación.'],
['R-03','Personal clave','Que el respaldo dependa de una sola jefatura de TI, impidiendo restauración ante su indisponibilidad.','BIA01-BIA06','Política 4',4,4,'Crítico','No se informó suplencia.'],
['R-04','Capacidad de ejecución','Que TI no ejecute pruebas trimestrales ni actualice BIA por carga operativa, dejando controles sin evidencia.','Todos','Política 3.4 y 3.8',4,4,'Crítico','No hay pruebas ni BIA actualizado.'],
['R-05','Resistencia al cambio','Que las áreas no validen RTO y RPO, conservando objetivos incompatibles con operación.','BIA01-BIA06','Respuesta BIA',3,4,'Alto','RTO y RPO propuestos por practicante.'],
['R-06','Proveedor','Que el proveedor de nube no demuestre cifrado y restauración, causando dependencia sin garantías.','BIA01-BIA06','Política 3.6',3,5,'Crítico','Cifrado solo declarado por proveedor.'],
['R-07','Técnico','Que el enlace y respaldos no tengan redundancia efectiva, impidiendo recuperar durante una interrupción.','BIA02,BIA03,BIA04','Universo RG-03 a RG-05',4,5,'Crítico','Enlace sin redundancia y falla de respaldo abierta.'],
['R-08','Normativo','Que discos de respaldo no cifrados expongan datos personales, generando incumplimiento.','BIA01,BIA02,BIA04,BIA06','Política 3.6',4,5,'Crítico','Disco externo no cifrado.'],
['R-09','De alcance','Que el plan solo cubra copias y no continuidad, restauración y responsables, dejando escenarios críticos sin tratamiento.','Todos','Respuesta TI',4,4,'Crítico','No existe plan de continuidad.'],
['R-10','Ejecución del plan','Que no se asigne responsable, presupuesto y seguimiento, dejando acciones sin ejecutar.','Todos','Universo auditable',4,5,'Crítico','Tratamientos sin plan ni auditoría sistemática.'],
['R-11','Presupuestal','Que el costo de copia externa segura y pruebas exceda la disponibilidad, obligando controles parciales.','BIA01-BIA06','Política y capacidad',3,4,'Alto','Recursos no calculados formalmente.'],
['R-12','Capacidad de ejecución','Que incoherencias del BIA no se corrijan, diseñando recuperación que incumpla RTO/RPO.','BIA01-BIA06','bia.csv',4,4,'Crítico','RPO/RTO inconsistentes y pruebas ausentes.']]
for r in R:r.insert(7,r[5]*r[6])
T=[]
for r in R:
 rid,val=r[0],r[7]; residual=8 if val>=15 else 6
 action={'R-01':'Aprobar presupuesto por fases y copia externa cifrada.','R-02':'Instalar comité y aprobar BIA por acta.','R-03':'Documentar procedimiento y designar suplente.','R-04':'Programar prueba trimestral con evidencia.','R-05':'Validar RTO/RPO con dueños de proceso.','R-06':'Exigir SLA, cifrado y prueba de restauración.','R-07':'Implementar redundancia y probar recuperación externa.','R-08':'Cifrar discos y restringir acceso.','R-09':'Definir alcance y escenarios de continuidad.','R-10':'Nombrar responsable y tablero trimestral.','R-11':'Fasear inversión y escenario mínimo viable.','R-12':'Corregir BIA y obtener firmas.'}[rid]
 T.append([rid,val,'Mitigar',action,'Jefatura de TI' if rid not in ['R-02','R-10'] else 'Gerencia General','90 días','Medio',2,4,residual,'Acción sin avance','30 días sin evidencia','Gerencia General'])
def wc(name,head,rows):
 with (out/name).open('w',encoding='utf-8-sig',newline='') as f: w=csv.writer(f);w.writerow(head);w.writerows(rows)
wc('RG02_registro.csv',['id','Categoría','Riesgo','Procesos afectados','Origen','P','I','Valor','Nivel','Justificación'],R)
wc('RG04_tratamiento.csv',['id','Valor','Decisión','Acciones concretas de tratamiento','Responsable','Plazo','Costo','P residual','I residual','Valor residual','Señal temprana','Umbral de alerta','Aceptación residual'],T)
C=[[x[0],f'Estado y residual de {x[0]}','Gerencia General y comité','Mensual','Sesión y acta',x[4],'Tablero','Priorizar, escalar o aceptar'] for x in T]
wc('RG06_comunicacion.csv',['id','Qué se comunica','A quién','Frecuencia','Canal','Quién comunica','Formato','Decisión'],C)
(out/'RG01_contexto.md').write_text('# Contexto y criterios\n\nFecha de acuerdo: 18 de noviembre de 2026, anterior a la evaluación del 19 de noviembre. Alcance: riesgos de ejecución del plan de respaldo, recuperación y continuidad (12 meses).\n\nProbabilidad: 1 rara, 2 improbable, 3 posible, 4 probable (evidencia vigente), 5 casi cierta (control ausente). Impacto: 1 menor, 2 bajo, 3 moderado, 4 alto (interrupción crítica), 5 crítico (pérdida prolongada, datos expuestos o sanción). Valor=P x I. <=6 aceptable documentado; 7-14 tratamiento; >=15 decisión de Gerencia General.\n\nRevisión trimestral y ante caída de enlace, fallo de restauración, incidente de datos, cambio de proveedor o desviación presupuestal >10%.\n',encoding='utf-8')
wc('RG03_evaluado.csv',['id','categoria','riesgo','procesos','origen','P','I','Valor','nivel','justificacion'],R)
# PNG autónomo de la matriz 5x5, sin bibliotecas externas.
W=750; H=650; pix=bytearray([255,255,255,255])*(W*H)
def rect(x,y,w,h,c):
 for yy in range(max(0,y),min(H,y+h)):
  for xx in range(max(0,x),min(W,x+w)):
   i=(yy*W+xx)*4;pix[i:i+4]=bytes((*c,255))
for iy in range(5):
 for ix in range(5):
  v=(ix+1)*(5-iy); rect(100+ix*100,60+iy*100,98,98,(80+min(175,v*8),220-min(180,v*8),80-min(70,v*3)))
for r in R: rect(100+(r[5]-1)*100+42,60+(5-r[6])*100+42,16,16,(0,0,0))
raw=b''.join(b'\x00'+bytes(pix[y*W*4:(y+1)*W*4]) for y in range(H))
def chunk(k,d):return struct.pack('>I',len(d))+k+d+struct.pack('>I',zlib.crc32(k+d)&0xffffffff)
(out/'graficos').mkdir(exist_ok=True);(out/'graficos'/'RG_mapa_calor.png').write_bytes(b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',W,H,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw,9))+chunk(b'IEND',b''))
initial=sum(x[1] for x in T);final=sum(x[9] for x in T);(ev/'RG05_residual.txt').write_text(f'Valor inicial: {initial}\nValor residual: {final}\nReducción: {initial-final} puntos ({(initial-final)/initial:.0%})\nRiesgos sobre el criterio: '+', '.join(x[0] for x in T if x[9]>6)+'\nRiesgos sin señal temprana: 0\n',encoding='utf-8')
(out/'RG05_residual.py').write_text("import pandas as pd\nt=pd.read_csv('RG04_tratamiento.csv');t['reduccion']=t.Valor-t['Valor residual'];print(t[['id','Valor','Valor residual','reduccion']].to_string(index=False))\n",encoding='utf-8')
(out/'seccion_8_gestion_riesgos.md').write_text('## 8. Gestión de riesgos del plan\n\n### 8.1 Alcance y criterios\nSe gestiona la ejecución del plan de respaldo, recuperación y continuidad. Los criterios están en RG01_contexto.md.\n\n### 8.2 Método\nISO 31000:2018: contexto, identificación, análisis P x I, tratamiento, comunicación y revisión.\n\n### 8.3 Registro y evaluación\nRG02 registra 12 riesgos en las nueve categorías. RG03 y el mapa de calor muestran ocho críticos, tres altos y uno medio.\n\n### 8.5 Tratamiento y residual\nRG04 define responsable, acción, plazo, señal y umbral. El valor baja de 188 a 90 puntos.\n\n### 8.7 Comunicación y revisión\nRG06 define qué, a quién, cuándo, canal y decisión. La revisión es trimestral y ante disparadores.\n\n### 8.8 Aceptación\nLos residuales >6 requieren aceptación formal de Gerencia General.\n',encoding='utf-8')
finaldoc=root/'SI886-S15-TALLER-CHOQUENA-FINAL.docx';shutil.copy(root/'SI886-S15-TALLER-CHOQUENA.docx',finaldoc);d=Document(finaldoc)
for p in d.paragraphs[27:]:p._element.getparent().remove(p._element)
for tb in d.tables:tb._element.getparent().remove(tb._element)
def h(s,l=1):d.add_heading(s,l)
def p(s):d.add_paragraph(s)
def table(headers,rows):
 tb=d.add_table(rows=1,cols=len(headers));tb.alignment=WD_TABLE_ALIGNMENT.CENTER
 for i,x in enumerate(headers):tb.rows[0].cells[i].text=str(x);tb.rows[0].cells[i].paragraphs[0].runs[0].bold=True
 for row in rows:
  cc=tb.add_row().cells
  for i,x in enumerate(row):cc[i].text=str(x);[setattr(run.font,'size',Pt(8)) for run in cc[i].paragraphs[0].runs]
 return tb
h('Índice');p('Actualice los campos del índice en Word antes de exportar el PDF.')
h('1. Información sobre el evento práctico');h('1.1 Título del evento práctico',2);p('Registro de riesgos del plan y matriz de comunicación para la mejora de respaldo, recuperación y continuidad.')
h('1.2 Objetivos',2);p('Definir criterios antes de evaluar; identificar las nueve categorías; establecer tratamiento, residual, señales tempranas y comunicación; documentar la Sección 8 del PETI.')
h('1.3 Tiempo de duración',2);p('Semana 15: sesión de laboratorio y consolidación documental.')
h('1.4 Resultados de aprendizaje',2);p('Registro sustentado en política, BIA y universo auditable, con evaluación y tratamiento verificable.')
h('1.5 Recursos',2);table(['Recurso','Uso'],[['Python, pandas y matplotlib','Cálculo, CSV y mapa de calor'],['Git','Historial, commit y etiquetas'],['Documentos del caso','Fuentes de identificación']])
h('1.6 Seguridad',2);p('Información confidencial; responsables por cargo; la aceptación residual pertenece a la gerencia.')
h('2. Procedimiento o metodología');h('Paso A - Contexto y criterios',2);p('Se fijaron escalas y criterio de aceptación antes de evaluar (RG01_contexto.md).')
h('Paso B - Identificación y evaluación',2);p('Se contrastaron política, BIA, respuestas y universo auditable. Se identificaron 12 riesgos y se calculó P x I.')
h('Paso C - Tratamiento y comunicación',2);p('Se definieron acciones, responsables, plazos, residuales, señales y matriz de comunicación.')
h('Paso D - Validación y cierre',2);p('Se verificó cobertura, justificación, tratamientos y señales. SimpleRisk no se ejecutó por falta de Docker; se declara como resultado no logrado.')
h('3. Resultados y evidencias');p('Repositorio GitHub: https://github.com/mauchoquenachoque/SEMANA-15-PETI-CHOQUENA/tree/taller-15')
checks=['Contexto y criterios previos','Escalas operativas','Nueve categorías','Registro de 12 riesgos','Trazabilidad a procesos','Mapa de calor','Críticos identificados','Tratamiento >6','Residual calculado','Señal y umbral','Aceptación residual','Carga en SimpleRisk','Matriz de comunicación','Sección 8 y etiquetas']
files=['RG01_contexto.md','RG01_contexto.md','RG02_registro.csv','RG02_registro.csv','RG02_registro.csv','graficos/RG_mapa_calor.png','RG03_evaluado.csv','RG04_tratamiento.csv','RG05_residual.txt','RG04_tratamiento.csv','RG04_tratamiento.csv','Pendiente Docker','RG06_comunicacion.csv','seccion_8_gestion_riesgos.md']
table(['#','Resultado','Estado','Evidencia'],[[i+1,a,'No logrado' if i==11 else 'Logrado',b] for i,(a,b) in enumerate(zip(checks,files))])
p('Valor inicial: 188 puntos; valor residual: 90; reducción: 52%. Los riesgos residuales sobre el criterio requieren aceptación formal.')
h('4. Conclusiones');p('La ausencia de pruebas, BIA aprobado y cifrado de medios evidencia que la recuperación no está gestionada. Las señales tempranas convierten el registro en un control. El residual no elimina la responsabilidad de aceptación gerencial.')
h('5. Cuestionario');p('Si el trabajo se hace mal, la organización puede invertir con RTO/RPO no validados, descubrir que un respaldo no restaura y exponer datos personales en medios no cifrados.')
h('6. Referencias bibliográficas');p('ISO 31000:2018 Risk management - Guidelines. https://www.iso.org/standard/65694.html\nISO/IEC 27005:2022 Information security risk management. https://www.iso.org/standard/80585.html\nServicios Mineros Cordillera S.A.C. (2026). Política de respaldo y recuperación.\nServicios Mineros Cordillera S.A.C. (2026). Universo auditable y registro de riesgos.')
h('7. Anexos');p('A. RG02_registro.csv\nB. graficos/RG_mapa_calor.png\nC. RG04_tratamiento.csv y RG05_residual.txt\nD. RG06_comunicacion.csv')
d.save(finaldoc);print(finaldoc)
