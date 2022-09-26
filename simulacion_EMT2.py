import sys

#ruta del interprete de python que estamos usando
#ruta_interprete = 'C:\\XXXXXX\\Python38\\Lib\\site-packages\\'
#ruta_resultados = 'C:\\XXXXXX\\resultados\\'
ruta_interprete = 'C:\\Users\\56965\\Documents\\python\\digsilent\\interprete\\Python38\\Lib\\site-packages\\'
ruta_resultados = 'C:\\Users\\56965\\Documents\\python\\digsilent\\resultados\\'
sys.path.append(ruta_interprete)

#se importa libreria de power factory y del objeto COM para crear objetos de Test Universe de Omicron
import powerfactory as pf
import win32com.client

#Se crea el objeto Control Center del software Test Universe de Omicron
occApp = win32com.client.Dispatch("OCCenter.Application")
occApp.Visible = True
occDoc = occApp.Documents.Add()

#Se definen las caracteristicas de las fallas que se simularan sobre la linea
tipoFalla = {0:'3f',1:'2f', 2:'1f', 3:'2fg'}
resistencia = [0, 5, 10, 25, 50]
distancia = [1, 50, 99]

#Se definen los tiempos de simulacion, inicio, despeje y apertura de interruptores
t_ini = 1.5
t_despeje = 3
t_abrir = 3
t_stop = 4

#Se crea una instancia de digsilent
app = pf.GetApplication()
#Se crea una instancia de proyecto activo
actPr = app.GetActiveProject()
#Se crea una instancia del propio script
scr = app.GetCurrentScript()
#Se crea una instancia del caso de esudio activo
stCase = app.GetActiveStudyCase()
#Se crea una instancia del comando de exportacion de resultados
exportResults = scr.GetContents('ExportComtrade.ComRes')[0]
#Se crea un instancia de objeto de resultados
resultados = stCase.GetContents('Resultado.ElmRes')
#Se crea una instancia del caso de eventos
eventos = stCase.GetContents('eventos.IntEvt')

if len(resultados) > 0:
    for r in resultados:
        r.Delete()
        
resultados = None
resultado = stCase.CreateObject('ElmRes', 'Resultado')

#Se extraen los elementos TTMM
extremo1 = scr.extremo1
extremo1 = extremo1.All()

extremo2 = scr.extremo2
extremo2 = extremo2.All()

medidas = extremo1 + extremo2

#Variables de TTMM
varTC=['s:I2r_A','s:I2r_B','s:I2r_C']
varTP=['s:U2r_A','s:U2r_B','s:U2r_C']

#Se agregan variables al archivo de resultados
for m in medidas:
    if m.GetClassName() == 'StaVt':
        for v in varTP:
            resultado.AddVariable(m,str(v))
    if m.GetClassName() == 'StaCt':
        for v in varTC:
            resultado.AddVariable(m,str(v))

eventos = stCase.GetContents('eventos.IntEvt')

#Si existe el objeto de evento se elimina
if len(eventos) > 0:
    for e in eventos:
        e.Delete()
        
#Se crean los eventos de cortocircuito, despeje de falla y apertura de los interruptores
evento = stCase.CreateObject('IntEvt', 'eventos')

lineas = scr.lineas
lineas = lineas.All()
if lineas[0].GetClassName() == 'ElmLne':
    linea = lineas[0]

evtCC = evento.CreateObject('EvtShc', 'cortocircuito')
evtDes = evento.CreateObject('EvtShc', 'despejar')
evtOpen = evento.CreateObject('EvtSwitch', 'abrir')

#Se hacen ajustes de los eventos
#Distancia de linea
linea.ishclne = 1
#Se ajustan los parametros del evento de cortocircuito
evtCC.p_target = linea
evtCC.htime = 0
evtCC.mtime = 0
evtCC.time = t_ini
evtCC.X_f = 0

#Se ajustan los parametros del evento de despeje de falla
evtDes.i_shc = 4
evtDes.p_target = linea
evtDes.time = t_despeje
evtDes.htime = 0
evtDes.mtime = 0
evtDes.i_clearShc = 0

#Se ajustan los parametros del evento de apertura de falla
evtOpen.p_target = linea
evtOpen.htime = 0
evtOpen.mtime = 0
evtOpen.time = t_abrir
evtOpen.i_switch = 0
evtOpen.i_allph = 1

#Se crean comandos de condiciones iniciales y de simulacion EMT
calIni = app.GetFromStudyCase('ComInc')
runSim = app.GetFromStudyCase('ComSim')

calIni.iopt_sim = 'ins'
calIni.p_resvar = resultado
calIni.p_event = evento

runSim.tstop = t_stop

#Se crean listas de elementos TTMM de los extremos
tp1 = [o for o in extremo1 if o.GetClassName() == 'StaVt'][0]
tc1 = [o for o in extremo1 if o.GetClassName() == 'StaCt'][0]

tp2 = [o for o in extremo2 if o.GetClassName() == 'StaVt'][0]
tc2 = [o for o in extremo2 if o.GetClassName() == 'StaCt'][0]

#Se ajustan los parametros del comando de exportancion para generar archivos comtrade
exportResults.dSampling = 1000
exportResults.iopt_csel = 1
exportResults.pResult = resultado
exportResults.iopt_exp = 3

# hacer for para recorrer la distancia de falla, resistencia de falla y tipo de falla
app.ResetCalculation()

for falla in tipoFalla.keys():
    for d in distancia:
        for RF in resistencia:
            #Se ajustan los parametros resistencia de falla, tipo de falla y distancia de falla
            evtCC.R_f = RF
            evtCC.i_shc = falla
            linea.fshcloc = d
            
            #Se ejecutan los comandos de condiciones iniciales y de simulacion EMT
            calIni.Execute()
            runSim.Execute()
            
            #Se ajustan parametros del comando de exportacion y se ejecuta (exportar)
            exportResults.f_name = f"{ruta_resultados}resultado_{tipoFalla[falla]}_{d}_{RF}.dat"
            exportResults.resultobj = [resultado, resultado, resultado, resultado, resultado, resultado]
            exportResults.element = [tp1, tp1, tp1, tc1, tc1, tc1]
            exportResults.cvariable = varTP + varTC
            exportResults.Execute()
            
            #Se van creando los objetos Advanced TransPlay del objeto Control Center del software Test Universe de Omicron
            occTrans = occDoc.InsertObject("OMTrans.Document")
            occTrans.Name = f"resultado_{tipoFalla[falla]}_{d}_{RF}"
            occTransSpe = occTrans.Specific
            occTransDoc = occTransSpe.Document
            #Se inportan los archivos comtrade previamentes creados a partir de la simulacion
            occTransDoc.Import(f"{ruta_resultados}resultado_{tipoFalla[falla]}_{d}_{RF}.cfg",False)
            occTransSpe.Quit()
            occTransDoc = None
            occTransSpe = None
            occTrans = None
            
occDoc.SaveAs(f'{ruta_resultados}\\prueba_comtrade.occ')
occApp.Quit()