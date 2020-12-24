from reportlab.pdfgen import canvas
from reportlab.platypus import (SimpleDocTemplate, Paragraph, PageBreak, Image, Spacer, Table, TableStyle)
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.pagesizes import LETTER, inch
from reportlab.graphics.shapes import Line, LineShape, Drawing
from reportlab.lib.colors import Color
import os
import datetime 


class FooterCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []
        self.width, self.height = LETTER

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            if (self._pageNumber > 1):
                self.draw_canvas(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_canvas(self, page_count):
        page = "Página %s de %s" % (self._pageNumber, page_count)
        x = 350
        self.saveState()
        self.setStrokeColorRGB(0, 0, 0)
        self.setLineWidth(0.5)
        self.drawImage("NUEVO_LOGO.png", self.width-inch*8-5, self.height-50, width=100, height=20, preserveAspectRatio=True)
        self.drawImage("Bandera.JPG", self.width - inch * 2, self.height-50, width=100, height=30, preserveAspectRatio=True, mask='auto')
        self.line(30, 740, LETTER[0] - 50, 740)         # line(Surface, color, (x1,y1), (x2,y2), width)
        self.line(66, 78, LETTER[0] - 66, 78)
        self.setFont('Times-Roman', 10)
        self.drawString(LETTER[0]-x, 65, page)
        self.restoreState()

class PDFPSReporte:

    def __init__(self, path):
        self.path = path
        self.styleSheet = getSampleStyleSheet()
        self.elements = []

        # colors - Azul turkeza 367AB3
        self.colorOhkaGreen0 = Color((45.0/255), (166.0/255), (153.0/255), 1)
        self.colorOhkaGreen1 = Color((182.0/255), (227.0/255), (166.0/255), 1)
        self.colorOhkaGreen2 = Color((140.0/255), (222.0/255), (192.0/255), 1)
        #self.colorOhkaGreen2 = Color((140.0/255), (222.0/255), (192.0/255), 1)
        self.colorOhkaBlue0 = Color((54.0/255), (122.0/255), (179.0/255), 1)
        self.colorOhkaBlue1 = Color((122.0/255), (180.0/255), (225.0/255), 1)
        self.colorOhkaGreenLineas = Color((50.0/255), (140.0/255), (140.0/255), 1)

        self.firstPage()
        self.nextPagesHeader(True)
        self.remoteSessionTableMaker()
        self.nextPagesHeader(False)
        self.inSiteSessionTableMaker()
        self.nextPagesHeader(False)
        self.extraActivitiesTableMaker()
        self.nextPagesHeader(False)
        self.summaryTableMaker()
        # Build
        self.doc = SimpleDocTemplate(path, pagesize=LETTER)
        self.doc.multiBuild(self.elements, canvasmaker=FooterCanvas)

    # Configura el diseño de la portada
    def firstPage(self):
        img = Image('NUEVO_LOGO.png', kind='proportional')
        img.drawHeight = 0.5*inch
        img.drawWidth = 2.4*inch
        img.hAlign = 'LEFT'
        self.elements.append(img)

        spacer = Spacer(30, 100)
        self.elements.append(spacer)
        
        psDetalle = ParagraphStyle('Resumen', fontSize=16, leading=14, justifyBreaks=1, alignment=TA_CENTER, justifyLastLine=1)
        text = """<strong>INFORME DE ENSAYO DE MOTOR</strong><br/><br/><br/>"""
        paragraphReportSummary = Paragraph(text, psDetalle)
        self.elements.append(paragraphReportSummary)
        spacer = Spacer(10, 100)
        
        img = Image('Portada.png')
        img.drawHeight = 4*inch
        img.drawWidth = 4*inch
        self.elements.append(img)

        spacer = Spacer(10, 100)
        self.elements.append(spacer)

        psDetalle = ParagraphStyle('Resumen', fontSize=9, leading=14, justifyBreaks=1, alignment=TA_LEFT, justifyLastLine=1)
        text = """INFORME DE ENSAYO MOTOR<br/>
        <strong>Empresa:</strong> COHEMO<br/>
        <strong>Fecha de Inicio:</strong> 23-Oct-2019<br/>
        <strong>Fecha de actualización:</strong> 01-Abril-2020<br/>
        """
        paragraphReportSummary = Paragraph(text, psDetalle)
        self.elements.append(paragraphReportSummary)
        self.elements.append(PageBreak())

    # Configura el diseño de las demás páginas
    def nextPagesHeader(self, isSecondPage):
        if isSecondPage:
            psHeaderText = ParagraphStyle('Hed0', fontSize=16, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaGreen0)
            text = 'Pruebas motores'
            paragraphReportHeader = Paragraph(text, psHeaderText)
            self.elements.append(paragraphReportHeader)

            spacer = Spacer(10, 10)
            self.elements.append(spacer)

            d = Drawing(500, 1)
            line = Line(-15, 0, 483, 0)
            line.strokeColor = self.colorOhkaGreenLineas
            line.strokeWidth = 2
            d.add(line)
            self.elements.append(d)

            spacer = Spacer(10, 1)
            self.elements.append(spacer)

            d = Drawing(500, 1)
            line = Line(-15, 0, 483, 0)
            line.strokeColor = self.colorOhkaGreenLineas
            line.strokeWidth = 0.5
            d.add(line)
            self.elements.append(d)

            spacer = Spacer(10, 22)
            self.elements.append(spacer)

    def remoteSessionTableMaker(self):        
        psHeaderText = ParagraphStyle('Hed0', fontSize=12, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaBlue0)
        text = 'Conexión y desconexión de ventiladores'
        paragraphReportHeader = Paragraph(text, psHeaderText)
        self.elements.append(paragraphReportHeader)

        spacer = Spacer(10, 22)
        self.elements.append(spacer)
        """
        Create the line items
        """
        "Importar librerias necesarias"

        import pandas as pd
        import matplotlib.pyplot as plt
        import numpy as np
        import math 
        #from array import *

        "Ruta en la que se desea guardar los graficos"
        ruta=r"Z:\INTERCAMBIO DATOS\DAVID VIGÓN\GMP - Generador de informes\PDF_creator"

        "Extraer datos y declarar matrices"
        datos= r"Z:\INGENIERIA\02_PROYECTOS\2020_09_24_AUTOMATIZACION_BANCO_GMPs\03_PROYECTO\034_DISEÑO\0343_SOFTWARE\GraficasGMP\GraficasGMP\Recursos\Prueba_funcional_3784.txt"
        datos= pd.read_csv(datos,sep="\t")

        Ventilador1=datos['VENTILADOR1']
        Ventilador2=datos['VENTILADOR2']
        Revoluciones=datos['Revoluciones']
        Potencia=datos['KW+']
        Par=datos['Par']
        TEMP_MKE=datos['TEMP_MKE']
        Fases=datos['Pattern Phase']
        KICK_DOWN1=datos['KICK-DOWN1']
        KICK_DOWN2=datos['KICK-DOWN2']




        "Bloque 1: encontrar los puntos"
        'Puntos de temperatura' 

        Vent1_puntos=np.zeros([len(Ventilador1),2])#matriz de los puntos donde se enchufa y se apaga el ventilador
        Vent2_puntos=np.zeros([len(Ventilador2),2])#matriz de los puntos donde se enchufa y se apaga el ventilador
        i1=0
        i2=0
        for z in range(1, len(Ventilador1),1):
            if Ventilador1[z]!=Ventilador1[z-1] and (Fases[z]== 'F0- TABLA CONEX VENTIADORES'or Fases[z]=='F0- TABLA DESCONEX VENTIADORES'):#Ventilador1.diff() funcion diferencial
                Vent1_puntos[i1,0]=TEMP_MKE[z]
                Vent1_puntos[i1,1]=z        
                i1=i1+1
            if Ventilador2[z]!=Ventilador2[z-1] and (Fases[z]== 'F0- TABLA CONEX VENTIADORES'or Fases[z]=='F0- TABLA DESCONEX VENTIADORES'):
                Vent2_puntos[i2,0]=TEMP_MKE[z]
                Vent2_puntos[i2,1]=z        
                i2=i2+1
 

        'Puntos de Kick-down1 y Kick-down2 ' 
        Kick_down1_puntos=np.zeros([len(KICK_DOWN1),2])#matriz de los puntos donde se enchufa y se apaga el kick-down1
        Kick_down2_puntos=np.zeros([len(KICK_DOWN2),2])#matriz de los puntos donde se enchufa y se apaga el kick-down2
        i3=0
        i4=0
        for z in range(1, len(KICK_DOWN1),1):
            if KICK_DOWN1[z]!=KICK_DOWN1[z-1] and (Fases[z]== 'F1_APROXIMACION KD1'or Fases[z]=='F1_TABLA CONEX KICKDOWN 1'):
                Kick_down1_puntos[i3,0]=Par[z]
                Kick_down1_puntos[i3,1]=z        
                i3=i3+1
            if KICK_DOWN2[z]!=KICK_DOWN2[z-1] and (Fases[z]== 'F2_APROXIMACION KD2'or Fases[z]=='F2_TABLA CONEX KICKDOWN 2'):
                Kick_down2_puntos[i4,0]=Par[z]
                Kick_down2_puntos[i4,1]=z        
                i4=i4+1
        
        
        'Puntos caida de potencia temperatura'  
        Potencia_puntos=np.zeros([len(KICK_DOWN2),3])#matriz de los puntos donde baja la potencia en funcion de la temperatura
        i5=0  
        for z in range(0, len(Potencia),1):
            if Fases[z]== 'F3_TABLA_PERD. POT':
                Potencia_puntos[i5,0]=Potencia[z]
                Potencia_puntos[i5,1]=z
                Potencia_puntos[i5,2]=i5         
                i5=i5+1
        
        #buscamos el máximo y mínimo y su posicion
        global Max_perdida
        global Min_perdida
        Max_perdida=np.zeros([1,3])
        Min_perdida=np.zeros([1,3])
          
        Max_perdida[0,0]=np.amax(Potencia_puntos[:i5,0])
        Min_perdida[0,0]=np.amin(Potencia_puntos[:i5,0])


        Max_perdida[0,1]=Potencia_puntos[np.argmax(Potencia_puntos[:i5,0]),2]      
        Min_perdida[0,1]=Potencia_puntos[np.argmin(Potencia_puntos[:i5,0]),2]

        Max_perdida[0,2]=Potencia_puntos[np.argmax(Potencia_puntos[:i5,0]),1]      
        Min_perdida[0,2]=Potencia_puntos[np.argmin(Potencia_puntos[:i5,0]),1]


        Temp_max_perdida=TEMP_MKE[Max_perdida[0,2]]
        Temp_min_perdida=TEMP_MKE[Min_perdida[0,2]]  



        "Rampa desconexión 1200rpm"
        Desconexion_1200=np.zeros(len(Fases))
        i6=0  
        for z in range(0, len(Potencia),1):
            if Fases[z]=='F7_TABLA_DESC. VENTILA 1200':
                Desconexion_1200[i6]=z        
                i6=i6+1


        "Rampa desconexión 1400rpm"
        Desconexion_1400=np.zeros(len(Fases))
        i7=0  
        for z in range(0, len(Potencia),1):
            if Fases[z]=='F9_TABLA_DESC. VENTILA 1400':
                Desconexion_1400[i7]=z        
                i7=i7+1
        
        
        "Prueba regulador"
        Regulador=np.zeros(len(Fases))
        i8=0
        for z in range(0, len(Potencia),1):
            if Fases[z]=='F5_PRB REGULADOR':
                Regulador[i8]=z        
                i8=i8+1        

        "Bloque 2: gráficos"
        'Conexión/desconexión ventiladores'
        y1=np.zeros(i1)
        y2=np.zeros(i2)

        Ventilador_graf= plt.figure(figsize=(15,10))

        ax0=plt.subplot(311)
        P=ax0.plot(np.linspace(0, len(Potencia), len(Potencia), endpoint=True), Potencia,'tab:cyan', label='Potenica real')
        plt.ylabel('Potencia (kW)')
        plt.xlabel('Número de muestreo')
        ax0.grid(True)

        ax1 = ax0.twinx()

        T=ax1.plot(np.linspace(0, len(Potencia), len(Potencia), endpoint=True), TEMP_MKE,'crimson', label='Temperatura')
        plt.ylabel('Temperatura (°C)')
        #plt.legend(loc='best')
        lns = T+P
        labs = [l.get_label() for l in lns]
        ax1.legend(lns, labs,bbox_to_anchor=(1,0.6,0,0.2),loc='upper right')



        plt.subplot(312)
        plt.plot(np.linspace(0, len(Ventilador1), len(Ventilador1), endpoint=True), Ventilador1,'orange', label='Ventilador 1')
        plt.plot(Vent1_puntos[0:i1,1],y1+1,'go',label='Con/des')


        plt.grid(True)
        plt.legend(loc='best')
        plt.ylabel('Estado ventilador 1')
        plt.xlabel('Número de muestreo')

        plt.subplot(313)
        plt.plot(np.linspace(0, len(Ventilador2), len(Ventilador2), endpoint=True), Ventilador2, label='Ventilador 2')
        plt.plot(Vent2_puntos[0:i2,1],y2+1,'ro',label='Con/des')
        plt.legend(loc='best')
        plt.ylabel('Estado ventilador 2')
        plt.xlabel('Número de muestreo')
        plt.grid(True)



        columnas=('Temperatura Conexión','Temperatura desconexión')
        filas=('Ventilador 1','Ventilador 2')
        n_filas=i1
        y_offset = np.zeros(len(columnas))
        cellText = [[Vent1_puntos[0,0],Vent1_puntos[1,0]],[Vent2_puntos[0,0],Vent2_puntos[1,0]]]
        tabla = plt.table(cellText=cellText, rowLoc='right',
                 rowLabels=filas,
                 colWidths=[.5,.5], colLabels=columnas,
                 colLoc='center', loc='bottom',bbox=[0,-0.8,1,0.5],zorder=20)

        tabla.auto_set_font_size(True)
        tabla.scale(1,1)
        plt.grid(True)
        ax0.set_title('Conexión/desconexión de ventiladores', fontsize=16)

        plt.tight_layout()
        plt.savefig('Ventilador_graf.png', dpi = 108)
        img = Image('Ventilador_graf.png')
        img.drawHeight = 540
        img.drawWidth = 540*1.1
        self.elements.append(img)
        
        'Conexión Kick-Down'
        y3=np.zeros(i3)
        y4=np.zeros(i4)
        
        Kic_down_graf, ax= plt.subplots(figsize=(15,10))
        ax=plt.subplot(311)
        P=ax.plot(np.linspace(0, len(Potencia), len(Potencia), endpoint=True), Potencia,'tab:cyan', label='Potencia')
        plt.ylabel('Potencia (kW)')
        ax1 = ax.twinx()
        
        T=ax1.plot(np.linspace(0, len(Par), len(Par), endpoint=True), Par,'crimson', label='Par')
        lns = T+P
        labs = [l.get_label() for l in lns]
        ax1.legend(lns, labs,bbox_to_anchor=(1,0.6,0,0.2),loc='upper right')
        plt.ylabel('Par (N·m)')
        plt.xlabel('Número de muestreo')
        plt.grid(True)
        plt.xlabel('Número de muestreo')
        Kic_down_graf.suptitle('Conexión Kick-Down', y=1.002, fontsize=16)
        
        
        plt.subplot(312)
        plt.plot(np.linspace(0, len(KICK_DOWN1), len(KICK_DOWN1), endpoint=True), KICK_DOWN1,'orange', label='KICK_DOWN1 1')
        plt.plot(Kick_down1_puntos[0:i3,1],y3+1,'go',label='Punto de conexión')
        
        
        plt.grid(True)
        plt.legend(loc='best')
        plt.ylabel('Estado Kick-Down 1')
        plt.xlabel('Número de muestreo')
        
        plt.subplot(313)
        plt.plot(np.linspace(0, len(KICK_DOWN2), len(KICK_DOWN2), endpoint=True), KICK_DOWN2,'b', label='KICK_DOWN2')
        plt.plot(Kick_down2_puntos[0:i4,1],y4+1,'go',label='Punto de conexión')
        plt.legend(loc='best')
        plt.ylabel('Estado Kick-Down 2')
        plt.xlabel('Número de muestreo')
        
        columnas=('Par (N·m)','Potencia (kW)')
        filas=('Kick-Down 1','Kick-Down 2')
        n_filas=i4
        y_offset = np.zeros(len(columnas))
        cellText = [[int(Kick_down1_puntos[0,0]),int(Potencia[Kick_down1_puntos[0,1]])],[int(Kick_down2_puntos[0,0]),int(Potencia[Kick_down2_puntos[0,1]])]]
        tabla = plt.table(cellText=cellText, rowLoc='right',
                          rowLabels=filas,
                          colWidths=[.5,.5], colLabels=columnas,
                           colLoc='center', loc='bottom',bbox=[0,-0.8,1,0.5],zorder=20)
        
        tabla.auto_set_font_size(True)
        tabla.scale(1,1)
        plt.grid(True)
        
        plt.grid(True)
        plt.tight_layout()
        plt.savefig("Kick_down_graf.png", dpi = 108)
        img = Image('Kick_down_graf.png')
        img.drawHeight = 540
        img.drawWidth = 540*1.1
        self.elements.append(img)
        
        
        
        #%%
        'Perdida potencia-temperatura'
        Perdida_potencia ,ax1= plt.subplots(figsize=(15,10))
        P=ax1.plot(Potencia_puntos[:i5,2], Potencia_puntos[:i5,0],'orange', label='Potencia real')
        plt.ylabel('Potencia(kW)')
        plt.xlabel('Número de muestra')
        plt.plot(Max_perdida[0,1],Max_perdida[0,0],'go',label='Máximo')
        plt.plot(Min_perdida[0,1],Min_perdida[0,0],'ro',label='Mínimo')
        
        ax2 = ax1.twinx()
        
        T=ax2.plot(Potencia_puntos[:i5,2], TEMP_MKE[Potencia_puntos[:i5,1]],'b', label='Temperatura MKE')
        plt.ylabel('Temperatura ºC')
        lns = T+P
        labs = [l.get_label() for l in lns]
        Perdida_potencia.legend(lns, labs,bbox_to_anchor=(0.3,0.6,0,0.2))
        ax1.yaxis.grid() # horizontal lines
        ax1.xaxis.grid() # vertical lines 
        columnas=('Potencia kW','Temperatura MKE ºC')
        filas=('Máximo valor','VenMínimo valor')
        n_filas=2
        y_offset = np.zeros(len(columnas))
        cellText = [[round(Max_perdida[0,0],2),round(Temp_max_perdida,1)],[round(Min_perdida[0,0],2),round(Temp_min_perdida,1)]]
        tabla = plt.table(cellText=cellText, rowLoc='right',
                         rowLabels=filas,
                         colWidths=[.5,.5], colLabels=columnas,
                         colLoc='center', loc='bottom',bbox=[0,-0.6,1,0.4],zorder=20)
        
        tabla.auto_set_font_size(True)
        tabla.scale(1,1)
        plt.title('Pérdida de potencia', fontsize=16)
        plt.tight_layout()
        plt.savefig("Perdida_potencia.png", dpi = 108)
        img = Image('Perdida_potencia.png')
        img.drawHeight = 540
        img.drawWidth = 540*1.1
        self.elements.append(img)
         
        
        #%%
        'Gráfico desconexión 1200 rpm' 
        
        Des_1200_graf, ax= plt.subplots(figsize=(15,10))
        plt.subplot(311)
        plt.plot(Revoluciones[Desconexion_1200[0:i6]],'tab:cyan', label='RPM motor')
        plt.legend(loc='best')
        plt.ylabel('Revoluciones/minuto')
        plt.xlabel('Número de muestreo')
        plt.grid(True)
        
        plt.subplot(312)
        plt.plot(Ventilador1[Desconexion_1200[0:i6]],'r', label='Ventilador 1')
        plt.plot(Ventilador2[Desconexion_1200[0:i6]],'b--', label='Ventilador 2')
        
        
        
        plt.grid(True)
        plt.legend(loc='best')
        plt.ylabel('Estado Ventiladores')
        plt.xlabel('Número de muestreo')
        
        plt.subplot(313)
        plt.plot(KICK_DOWN1[Desconexion_1200[0:i6]],'r', label='KICK_DOWN1')
        plt.plot(KICK_DOWN1[Desconexion_1200[0:i6]],'b--', label='KICK_DOWN2')
        plt.legend(loc='best')
        plt.ylabel('Estado kick-down')
        plt.xlabel('Número de muestreo')
        
        
        plt.grid(True)
        Des_1200_graf.suptitle('Desconexión 1200 rpm',y=1, fontsize=16)
        
        plt.grid(True)
        plt.tight_layout()
        plt.savefig("Des_1200_graf.png", dpi = 108)
        img = Image('Des_1200_graf.png')
        img.drawHeight = 540
        img.drawWidth = 540*0.9
        self.elements.append(img)
        
        
        #%%
        'Gráfico desconexión 1400 rpm' 
        
        Des_1400_graf=plt.figure(figsize=(15,10))
        ax0=plt.subplot(311)
        ax0.plot(Revoluciones[Desconexion_1400[0:i7]],'tab:cyan', label='RPM motor')
        ax0.legend(loc='best')
        ax0.set_ylabel('Revoluciones/minuto')
        ax0.set_xlabel('Número de muestreo')
        plt.grid(True)
        
        ax1=plt.subplot(312)
        ax1.plot(Ventilador1[Desconexion_1400[0:i7]],'r', label='Ventilador 1')
        ax1.plot(Ventilador2[Desconexion_1400[0:i7]],'b--', label='Ventilador 2')
        ax1.set_ylim([0,1.05])
        
        
        
        plt.grid(True)
        ax1.legend(loc='lower right')
        ax1.set_ylabel('Estado ventiladores')
        
        ax1.set_xlabel('Número de muestreo')
        
        ax2=plt.subplot(313)
        ax2.plot(KICK_DOWN1[Desconexion_1400[0:i7]],'r', label='KICK_DOWN1')
        ax2.plot(KICK_DOWN1[Desconexion_1400[0:i7]],'b--', label='KICK_DOWN2')
        ax2.legend(loc='best')
        ax2.set_ylabel('Estado kick-down')
        ax2.set_xlabel('Número de muestreo')
        
        
        plt.grid(True)
        Des_1400_graf.suptitle('Desconexión 1400 rpm', fontsize=16)
        
        plt.grid(True)
        plt.tight_layout()
        plt.savefig("Des_1400_graf.png", dpi = 108)
        img = Image('Des_1400_graf.png')
        img.drawHeight = 540
        img.drawWidth = 540*0.9
        self.elements.append(img)
        
        
        #%%
        'Prueba regulador'
        def make_patch_spines_invisible(ax):
            ax.set_frame_on(True)
            ax.patch.set_visible(False)
            for sp in ax.spines.values():
                sp.set_visible(False)
        
        
        fig, host = plt.subplots()
        fig.subplots_adjust(right=0.75)
        
        par1 = host.twinx()
        par2 = host.twinx()
        
        # Offset the right spine of par2.  The ticks and label have already been
        # placed on the right by twinx above.
        par2.spines["right"].set_position(("axes", 1.2))
        # Having been created by twinx, par2 has its frame off, so the line of its
        # detached spine is invisible.  First, activate the frame but make the patch
        # and spines invisible.
        make_patch_spines_invisible(par2)
        # Second, show the right spine.
        par2.spines["right"].set_visible(True)
        
        p1, = host.plot(np.linspace(0, i8*5, i8),Potencia[Regulador[0:i8]],'r--', label='Potencia real')
        p2, = par1.plot(np.linspace(0, i8*5, i8),Par[Regulador[0:i8]],'b--', label='Par')
        p3, = par2.plot(np.linspace(0, i8*5, i8),Revoluciones[Regulador[0:i8]],'g', label='Revoluciones')
        
        
        
        host.set_xlabel("Tiempo (s)")
        host.set_ylabel("Potencia (kW)")
        par1.set_ylabel("Par (Nm)")
        par2.set_ylabel("Revoluciones (rpm)")
        
        
        
        tkw = dict(size=4, width=1.5)
        host.tick_params(axis='y', colors=p1.get_color(), **tkw)
        par1.tick_params(axis='y', colors=p2.get_color(), **tkw)
        par2.tick_params(axis='y', colors=p3.get_color(), **tkw)
        
        
        
        lines = [p1, p2, p3]
        
        host.legend(lines, [l.get_label() for l in lines],loc='center right')
        
        plt.tight_layout()
        plt.title('Regulador 2680',fontsize=16)
        plt.grid(True)
        plt.savefig("Regulador_2680.png", dpi = 108)
        img = Image('Regulador_2680.png')
        img.drawHeight = 500*0.9
        img.drawWidth = 500
        self.elements.append(img)



        
        """self.drawImage("Ventilador_graf.png", 
                       self.width-inch*8-5, 
                       self.height-50, 
                       width=100, 
                       height=20, 
                       preserveAspectRatio=True)"""
        #canvas.drawImage("NUEVO_LOGO.png", self.width - inch * 2, self.height-50, width=100, height=30, preserveAspectRatio=True, mask='auto')
        
        # d = []
        # textData = ["No.", "Fecha", "Hora Inicio", "Hora Fin", "Tiempo Total"]
                
        # fontSize = 8
        # centered = ParagraphStyle(name="centered", alignment=TA_CENTER)
        # for text in textData:
        #     ptext = "<font size='%s'><b>%s</b></font>" % (fontSize, text)
        #     titlesTable = Paragraph(ptext, centered)
        #     d.append(titlesTable)        

        # data = [d]
        # lineNum = 1
        # formattedLineData = []

        # alignStyle = [ParagraphStyle(name="01", alignment=TA_CENTER),
        #               ParagraphStyle(name="02", alignment=TA_LEFT),
        #               ParagraphStyle(name="03", alignment=TA_CENTER),
        #               ParagraphStyle(name="04", alignment=TA_CENTER),
        #               ParagraphStyle(name="05", alignment=TA_CENTER)]

        # for row in range(10):
        #     lineData = [str(lineNum), "Miércoles, 11 de diciembre de 2019", 
        #                                     "17:30", "19:24", "1:54"]
        #     #data.append(lineData)
        #     columnNumber = 0
        #     for item in lineData:
        #         ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
        #         p = Paragraph(ptext, alignStyle[columnNumber])
        #         formattedLineData.append(p)
        #         columnNumber = columnNumber + 1
        #     data.append(formattedLineData)
        #     formattedLineData = []
            
        # # Row for total
        # totalRow = ["Total de Horas", "", "", "", "30:15"]
        # for item in totalRow:
        #     ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
        #     p = Paragraph(ptext, alignStyle[1])
        #     formattedLineData.append(p)
        # data.append(formattedLineData)
        
        # #print(data)
        # table = Table(data, colWidths=[50, 200, 80, 80, 80])
        # tStyle = TableStyle([ #('GRID',(0, 0), (-1, -1), 0.5, grey),
        #         ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        #         #('VALIGN', (0, 0), (-1, -1), 'TOP'),
        #         ("ALIGN", (1, 0), (1, -1), 'RIGHT'),
        #         ('LINEABOVE', (0, 0), (-1, -1), 1, self.colorOhkaBlue1),
        #         ('BACKGROUND',(0, 0), (-1, 0), self.colorOhkaGreenLineas),
        #         ('BACKGROUND',(0, -1),(-1, -1), self.colorOhkaBlue1),
        #         ('SPAN',(0,-1),(-2,-1))
        #         ])
        # table.setStyle(tStyle)
        #self.elements.append(Ventilador_graf.png)
          
    def inSiteSessionTableMaker(self):
        self.elements.append(PageBreak())
        psHeaderText = ParagraphStyle('Hed0', fontSize=12, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaBlue0)
        text = 'SESIONES EN SITIO'
        paragraphReportHeader = Paragraph(text, psHeaderText)
        self.elements.append(paragraphReportHeader)

        spacer = Spacer(10, 22)
        self.elements.append(spacer)
        """
        Create the line items
        """
        d = []
        textData = ["Dato", "Fecha", "Potencia (kW)", "Temperatura (°C)"]
                
        fontSize = 8
        centered = ParagraphStyle(name="centered", alignment=TA_CENTER)
        for text in textData:
            ptext = "<font size='%s'><b>%s</b></font>" % (fontSize, text)
            titlesTable = Paragraph(ptext, centered)
            d.append(titlesTable)        

        data = [d]
        lineNum = ["Máximo valor","Mínimo valor"]
        formattedLineData = []

        alignStyle = [ParagraphStyle(name="01", alignment=TA_CENTER),
                      ParagraphStyle(name="02", alignment=TA_LEFT),
                      ParagraphStyle(name="03", alignment=TA_CENTER),
                      ParagraphStyle(name="04", alignment=TA_CENTER),
                      ParagraphStyle(name="05", alignment=TA_CENTER)]

        for row in range(2):
            lineData = [str(lineNum[row]), datetime.datetime.now().strftime("%H:%M:%S %d/%m/%Y"), 
                                           round( Max_perdida[0,0],2),
                                           round(Min_perdida[0,0],2)]
            #data.append(lineData)
            columnNumber = 0
            for item in lineData:
                ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
                p = Paragraph(ptext, alignStyle[columnNumber])
                formattedLineData.append(p)
                columnNumber = columnNumber + 1
            data.append(formattedLineData)
            formattedLineData = []
            
        # Row for total
        # totalRow = ["Total de Horas", "", "", "", "30:15"]
        # for item in totalRow:
        #     ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
        #     p = Paragraph(ptext, alignStyle[1])
        #     formattedLineData.append(p)
        data.append(formattedLineData)
        
        print(data)
        table = Table(data, colWidths=[50, 200, 80, 80, 80])
        tStyle = TableStyle([ #('GRID',(0, 0), (-1, -1), 0.5, grey),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                #('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ("ALIGN", (1, 0), (1, -1), 'RIGHT'),
                ('LINEABOVE', (0, 0), (-1, -1), 1, self.colorOhkaBlue1),
                ('BACKGROUND',(0, 0), (-1, 0), self.colorOhkaGreenLineas),
                ('BACKGROUND',(0, -1),(-1, -1), self.colorOhkaBlue1),
                ('SPAN',(0,-1),(-2,-1))
                ])
        table.setStyle(tStyle)
        self.elements.append(table)

    def extraActivitiesTableMaker(self):
        self.elements.append(PageBreak())
        psHeaderText = ParagraphStyle('Hed0', fontSize=12, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaBlue0)
        text = 'OTRAS ACTIVIDADES Y DOCUMENTACIÓN'
        paragraphReportHeader = Paragraph(text, psHeaderText)
        self.elements.append(paragraphReportHeader)

        spacer = Spacer(10, 22)
        self.elements.append(spacer)
        """
        Create the line items
        """
        d = []
        textData = ["Dato.", "Fecha", "Potencia (kW)", "Temperatura (°C)"]
                
        fontSize = 8
        centered = ParagraphStyle(name="centered", alignment=TA_CENTER)
        for text in textData:
            ptext = "<font size='%s'><b>%s</b></font>" % (fontSize, text)
            titlesTable = Paragraph(ptext, centered)
            d.append(titlesTable)        

        data = [d]
        lineNum = 1
        formattedLineData = []

        alignStyle = [ParagraphStyle(name="01", alignment=TA_CENTER),
                      ParagraphStyle(name="02", alignment=TA_LEFT),
                      ParagraphStyle(name="03", alignment=TA_CENTER),
                      ParagraphStyle(name="04", alignment=TA_CENTER),
                      ParagraphStyle(name="05", alignment=TA_CENTER)]
       
        for row in range(2):
            lineData = [str(lineNum), datetime.datetime.now(), 
                                            round(Max_perdida[0,0],2), round(Min_perdida[0,0],2), "1:54"]
            #data.append(lineData)
            columnNumber = 0
            for item in lineData:
                ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
                p = Paragraph(ptext, alignStyle[columnNumber])
                formattedLineData.append(p)
                columnNumber = columnNumber + 1
            data.append(formattedLineData)
            formattedLineData = []
            
        # Row for total
        totalRow = [ "Total de Horas", "", "", "", "30:15"]
        for item in totalRow:
            ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
            p = Paragraph(ptext, alignStyle[1])
            formattedLineData.append(p)
        data.append(formattedLineData)
        
        print(data)
        table = Table(data, colWidths=[50, 200, 80, 80, 80])
        tStyle = TableStyle([ #('GRID',(0, 0), (-1, -1), 0.5, grey),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                #('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ("ALIGN", (1, 0), (1, -1), 'RIGHT'),
                ('LINEABOVE', (0, 0), (-1, -1), 1, self.colorOhkaBlue1),
                ('BACKGROUND',(0, 0), (-1, 0), self.colorOhkaGreenLineas),
                ('BACKGROUND',(0, -1),(-1, -1), self.colorOhkaBlue1),
                ('SPAN',(0,-1),(-2,-1))
                ])
        table.setStyle(tStyle)
        self.elements.append(table)
        
        spacer = Spacer(30, 100)
        psDetalle = ParagraphStyle('Resumen', fontSize=10, font='Arial', leading=14, justifyBreaks=1, alignment=TA_LEFT, justifyLastLine=1)
        text = """<br/><p>En la tabla de arriba se muestran los valores de potencia y temperatura en la perdida de potencia por temperatura</p><br/><br/><br/>"""
        paragraphReportSummary = Paragraph(text, psDetalle)
        self.elements.append(paragraphReportSummary)
        spacer = Spacer(10, 100)

    def summaryTableMaker(self):
        self.elements.append(PageBreak())
        psHeaderText = ParagraphStyle('Hed0', fontSize=12, alignment=TA_LEFT, borderWidth=3, textColor=self.colorOhkaBlue0)
        text = 'REGISTRO TOTAL DE HORAS'
        paragraphReportHeader = Paragraph(text, psHeaderText)
        self.elements.append(paragraphReportHeader)

        spacer = Spacer(10, 22)
        self.elements.append(spacer)
        """
        Create the line items
        """

        tStyle = TableStyle([
                    ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                    #('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ("ALIGN", (1, 0), (1, -1), 'RIGHT'),
                    ('LINEABOVE', (0, 0), (-1, -1), 1, self.colorOhkaBlue1),
                    ('BACKGROUND',(-2, -1),(-1, -1), self.colorOhkaGreen2)
                    ])

        fontSize = 8
        lineData = [["Sesiones remotas", "30:15"],
                    ["Sesiones en sitio", "00:00"],
                    ["Otras actividades", "00:00"],
                    ["Total de horas consumidas", "30:15"]]

        # for row in lineData:
        #     for item in row:
        #         ptext = "<font size='%s'>%s</font>" % (fontSize-1, item)
        #         p = Paragraph(ptext, centered)
        #         formattedLineData.append(p)
        #     data.append(formattedLineData)
        #     formattedLineData = []

        table = Table(lineData, colWidths=[400, 100])
        table.setStyle(tStyle)
        self.elements.append(table)

        # Total de horas contradas vs horas consumidas
        data = []
        formattedLineData = []

        lineData = [["Total de horas contratadas", "120:00"],
                    ["Horas restantes por consumir", "00:00"]]

        # for row in lineData:
        #     for item in row:
        #         ptext = "<b>{}</b>".format(item)
        #         p = Paragraph(ptext, self.styleSheet["BodyText"])
        #         formattedLineData.append(p)
        #     data.append(formattedLineData)
        #     formattedLineData = []

        table = Table(lineData, colWidths=[400, 100])
        tStyle = TableStyle([
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ("ALIGN", (1, 0), (1, -1), 'RIGHT'),
                ('BACKGROUND', (0, 0), (1, 0), self.colorOhkaBlue1),
                ('BACKGROUND', (0, 1), (1, 1), self.colorOhkaGreen1),
                ])
        table.setStyle(tStyle)

        spacer = Spacer(10, 50)
        self.elements.append(spacer)
        self.elements.append(table)

if __name__ == '__main__':

    nombre_archivo = 'Informe GMP__' + str( datetime.datetime.now().strftime("%d_%m_%Y___%H_%M_%S") )
    report = PDFPSReporte( nombre_archivo + '.pdf' )