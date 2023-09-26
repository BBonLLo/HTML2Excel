from tkinter import *
import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as tb
from ttkbootstrap.toast import ToastNotification
import pandas as pd
from bs4 import BeautifulSoup


def confirmar():
    if '.html' in lbl_file.cget('text'):
        # Estadisticas
        numero = []
        nombre = []
        partidosJugados = []
        minutosPartido = []
        puntos = []
        ppg = []
        tirosTotales = []
        tirosTotalesTirados = []
        tirosTotalesAnotados = []
        tirosTotalesPorcentage = []
        puntosDos = []
        puntosDosTirados = []
        puntosDosAnotados = []
        puntosDosPorcentage = []
        puntosTres = []
        puntosTresTirados = []
        puntosTresAnotados = []
        puntosTresPorcentage = []
        puntosLibres = []
        puntosLibresTirados = []
        puntosLibresAnotados = []
        puntosLibresPorcentage = []
        masmenos = []
        rebotes = []
        rebotesOfensivos = []
        rebotesDefensivos = []
        asistencias = []
        robos = []
        defl = []
        tapones = []
        taponesRecividos = []
        perdidas = []
        puntosEnPerdidas = []
        puntosEnSegundaOportunidad = []
        valoracion = []
        faltas = []
        faltasRecividas = []
        faltaEnAtaque = []
        ratioAsistenciasPerdidas = []

        try:
            # Ruta archivo html
            ruta = lbl_file.cget('text')
            if opcion.get()[0].lower() == 'p':
                equipo = opcion_equipo.get()[0].lower()
  
                if equipo == "l":
                  valor_equipo = 2
                else:
                  valor_equipo = 3
  
            else:
              valor_equipo = 0
            
            # Opening the html file
            HTMLFile = open(ruta, "r")
  
            # Cambiar terminacion para generar archivo excel
            ruta = ruta.replace("html", "xlsx", 1)
  
            # Reading the file
            index = HTMLFile.read()
  
            # Creating a BeautifulSoup object and specifying the parser
            S = BeautifulSoup(index, 'lxml')
  
            for tag in S.find_all('table')[valor_equipo]:
              todosJugadores = tag.text
              todosJugadoresSeparados = todosJugadores.split('\n')
              if "GAMES" not in todosJugadoresSeparados:
                  if "Totals" not in todosJugadoresSeparados:
                      if len(todosJugadoresSeparados) > 2:
                          new_jugador = todosJugadoresSeparados
  
                          # Printing the name, and text of p tag
                          new_jugador.pop(0)
                          new_jugador.pop(-1)
  
                          ultimoDato = new_jugador.pop(0)
                          # Separar numero de Jugador
                          if " " in ultimoDato:
                            ultimoDatoSeparado = ultimoDato.split(' ', 1)
                            numeroJugador = ultimoDatoSeparado.pop(0)
                            nombreJugador = ultimoDatoSeparado.pop(0)
  
                            if ("Ã¡" in nombreJugador or "Ã©" in nombreJugador or "Ã-­" in nombreJugador or "Ã³" in nombreJugador or "Ãº" in nombreJugador or "Ã" in nombreJugador or "Ã‰" in nombreJugador or "Ã" in nombreJugador or "Ã“" in nombreJugador or "Ãš" in nombreJugador or "Ã±" in nombreJugador):
                                nombreJugador = nombreJugador.replace("Ã¡", "á")
                                nombreJugador = nombreJugador.replace("Ã©", "é")
                                nombreJugador = nombreJugador.replace("­Ã-", "í")
                                nombreJugador = nombreJugador.replace("Ã³", "ó")
                                nombreJugador = nombreJugador.replace("Ãº", "ú")
                                nombreJugador = nombreJugador.replace("Ã", "Á")
                                nombreJugador = nombreJugador.replace("Ã‰", "É")
                                nombreJugador = nombreJugador.replace("Ã", "Í")
                                nombreJugador = nombreJugador.replace("Ã“", "Ó")
                                nombreJugador = nombreJugador.replace("Ãš", "Ú")
                                nombreJugador = nombreJugador.replace("Ã±", "ñ")
  
                          # Añadir datos a las listas
                          if " " in ultimoDato:
                            numero.append(numeroJugador)
                            nombre.append(nombreJugador)
                          else:
                            numero.append("#0/0")
                            nombre.append(ultimoDato)
  
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          partidosJugados.append(ultimoDato)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          minutosPartido.append(ultimoDato)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntos.append(ultimoDato)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          ppg.append(ultimoDato)
  
                          # Separar tirados y anotados (Tiros totales)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          tirosTotales.append(ultimoDato)
                          ultimoDatoSeparado = ultimoDato.split('/')
                          tirosAnotadosTotales = ultimoDatoSeparado.pop(0)
                          tirosIntentadosTotales = ultimoDatoSeparado.pop(0)
                          tirosTotalesAnotados.append(tirosAnotadosTotales)
                          tirosTotalesTirados.append(tirosIntentadosTotales)
  
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          tirosTotalesPorcentage.append(ultimoDato)
  
                          # Separar tirados y anotados (Tiros 2)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosDos.append(ultimoDato)
                          ultimoDatoSeparado = ultimoDato.split('/')
                          tirosDosAnotados = ultimoDatoSeparado.pop(0)
                          tirosDosIntentados = ultimoDatoSeparado.pop(0)
                          puntosDosAnotados.append(tirosDosAnotados)
                          puntosDosTirados.append(tirosDosIntentados)
  
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosDosPorcentage.append(ultimoDato)
  
                          # Separar tirados y anotados (Tiros 3)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosTres.append(ultimoDato)
                          ultimoDatoSeparado = ultimoDato.split('/')
                          tirosTresAnotados = ultimoDatoSeparado.pop(0)
                          tirosTresIntentados = ultimoDatoSeparado.pop(0)
                          puntosTresAnotados.append(tirosTresAnotados)
                          puntosTresTirados.append(tirosTresIntentados)
  
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosTresPorcentage.append(ultimoDato)
  
                          # Separar tirados y anotados (Tiros libres)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosLibres.append(ultimoDato)
                          ultimoDatoSeparado = ultimoDato.split('/')
                          tirosLibresAnotados = ultimoDatoSeparado.pop(0)
                          tirosLibresIntentados = ultimoDatoSeparado.pop(0)
                          puntosLibresAnotados.append(tirosLibresAnotados)
                          puntosLibresTirados.append(tirosLibresIntentados)
  
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosLibresPorcentage.append(ultimoDato)
  
                          # Masmenos
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          masmenos.append(ultimoDato)
  
                          # Rebotes
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          rebotes.append(ultimoDato)
  
                          # Rebotes ofensivos
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          rebotesOfensivos.append(ultimoDato)
  
                          # Rebotes defensivos
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          rebotesDefensivos.append(ultimoDato)
  
                          # Asistencias
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          asistencias.append(ultimoDato)
  
                          # Robos
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          robos.append(ultimoDato)
  
                          # DEFL (No se que es, Deflections By Team)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          defl.append(ultimoDato)
  
                          # Tapones
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          tapones.append(ultimoDato)
  
                          # Tapones recividos
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          taponesRecividos.append(ultimoDato)
  
                          # Perdidas
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          perdidas.append(ultimoDato)
  
                          # Puntos tras perdida
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosEnPerdidas.append(ultimoDato)
  
                          # Puntos en segunda oportunidad
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          puntosEnSegundaOportunidad.append(ultimoDato)
  
                          # Valoracion
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          valoracion.append(ultimoDato)
  
                          # Faltas
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          faltas.append(ultimoDato)
  
                          # Faltas recividas
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          faltasRecividas.append(ultimoDato)
  
                          # CT (Creo que faltas en ataque)
                          ultimoDato = todosJugadoresSeparados.pop(0)
                          faltaEnAtaque.append(ultimoDato)
  
                          # Asistencias por perdida
                          ultimoDato = todosJugadoresSeparados[0]
                          ratioAsistenciasPerdidas.append(ultimoDato)
  
            data = {
              "Numero": numero,
              "Nombre": nombre,
              "Partidos": partidosJugados,
              "Min": minutosPartido,
              "PTS": puntos,
              "PPP": ppg,
              "TC": tirosTotales,
              "%TC": tirosTotalesPorcentage,
              "2P": puntosDos,
              "%2P": puntosDosPorcentage,
              "3P": puntosTres,
              "%3P": puntosTresPorcentage,
              "TL": puntosLibres,
              "%TL": puntosLibresPorcentage,
              "+/-": masmenos,
              "REB": rebotes,
              "OREB": rebotesOfensivos,
              "DREB": rebotesDefensivos,
              "AST": asistencias,
              "STL": robos,
              "DEFL": defl,
              "BLK": tapones,
              "BLKR": taponesRecividos,
              "TO": perdidas,
              "POT": puntosEnPerdidas,
              "SCP": puntosEnSegundaOportunidad,
              "EFF": valoracion,
              "PF": faltas,
              "FOULR": faltasRecividas,
              "CT": faltaEnAtaque,
              "AST/TO": ratioAsistenciasPerdidas,
            }
  
            data2 = {
              "Nombre": nombre,
              "TCanot": tirosTotalesAnotados,
              "TCint": tirosTotalesTirados,
              "T2anot": puntosDosAnotados,
              "T2int": puntosDosTirados,
              "T3anot": puntosTresAnotados,
              "T3int": puntosTresTirados,
              "TLanot": puntosLibresAnotados,
              "TLint": puntosLibresTirados,
            }

            df = pd.DataFrame(data)
            df2 = pd.DataFrame(data2)

            df = df.astype({'Partidos':'int64', 'PTS':'int64', '+/-':'int64', 'REB':'int64', 'OREB':'int64',
                            'DREB':'int64', 'AST':'int64', 'STL':'int64', 'DEFL':'int64',
                            'BLK':'int64', 'BLKR':'int64', 'TO':'int64', 'POT':'int64',
                            'SCP':'int64', 'EFF':'int64', 'PF':'int64', 'FOULR':'int64',
                            'CT':'int64'
                          })

            df2 = df2.astype({'TCanot':'int64', 'TCint':'int64',
                              'T2anot':'int64', 'T2int':'int64',
                              'T3anot':'int64', 'T3int':'int64',
                              'TLanot':'int64', 'TLint':'int64'
                            })

            writer = pd.ExcelWriter(ruta, engine='xlsxwriter')
            workbook = writer.book
            worksheet = workbook.add_worksheet('Stats')
            writer.sheets['Stats'] = worksheet

            df.to_excel(writer, sheet_name = 'Stats', startcol = -1)
            df2.to_excel(writer, sheet_name = 'Stats', startrow = 20, startcol = -1)

            for i, col in enumerate(df.columns):
              width = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
              worksheet.set_column(i, i, width + 1.5)

            for i, col in enumerate(df2.columns):
              width = max(df2[col].apply(lambda x: len(str(x))).max(), len(col))
              worksheet.set_column(i, i, width + 1)

            worksheet.set_column(1, 1, 14)
            worksheet.set_column(2, 2, 9.86)
          
            writer.close()

            toast = ToastNotification(
                icon = '',
                title = 'HTML2Excel_V2.2.0',
                message = 'Todo perfecto inútil',
                duration = 5000,
                alert = True
            )
            toast.show_toast()
          
        except Exception:
            toast = ToastNotification(
                icon = '',
                title = 'HTML2Excel_V2.2.0',
                message = 'Algo ha fallado, contacta con Unai',
                duration = 5000,
                alert = True
            )
            
            toast.show_toast()
    else:
        lbl_file.config(text = 'Selecciona un archivo', foreground = 'red')

def salir():
    window.quit()

def borrarArchivo():
    if lbl_file.cget('text') == '':
        lbl_file.config(text = 'Selecciona un archivo', foreground = 'red')
    if lbl_file.cget('text') != 'Selecciona un archivo':
        lbl_file.config(text = '')
    
def buscarArchivo():
    filename = filedialog.askopenfilename(initialdir = '~/Downloads',
                                          title = 'Selecciona un archivo',
                                          filetypes = (('HTML files',
                                                        '*.html*'),
                                                       ('all files',
                                                        '*.*')))

    # Change label contents
    if len(filename) != 0:
        lbl_file.config(text = filename, foreground = '#366da3')

def click_partido():
    rb_local.configure(state = "enabled", cursor = 'hand2')
    rb_visitante.configure(state = "enabled", cursor = 'hand2')
    

def click_stats():
    rb_local.configure(state = "disabled", cursor = 'arrow')
    rb_visitante.configure(state = "disabled", cursor = 'arrow')

# Window
window = tb.Window(themename = 'journal', title = 'HTML2Excel_V2.2.0', resizable = (False, False))


app_width = 600
app_height = 350

screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - (app_height / 2)

window.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')

# Button styles
btn_style1 = tb.Style()
btn_style1.configure('info.TButton', font = ('Corbel', 10))

btn_style2 = tb.Style()
btn_style2.configure('info.TRadiobutton', font = ('Corbel', 10))

btn_style2_1 = tb.Style()
btn_style2_1.configure('info.Toolbutton', font = ('Corbel', 10))

btn_style3 = tb.Style()
btn_style3.configure('success.TButton', font = ('Corbel', 10))

btn_style4 = tb.Style()
btn_style4.configure('primary.TButton', font = ('Corbel', 10))

# Title
title_frame = tb.Frame(window, style = 'my.TFrame')
lbl_title = tb.Label(title_frame, text = 'HTML2Excel')
lbl_title.pack(side = 'left')
lbl_title.config(font = ('Corbel', 24, 'bold'), foreground = 'black')
title_frame.pack(fill = 'x', pady = 18, padx = 20)

# Insert file
insert_file_frame = tb.Frame(window, style = '#daeddb')
lbl_seleccionar_archivo = tb.Label(insert_file_frame, text = 'Selecciona un reporte: ', foreground = 'black')
lbl_seleccionar_archivo.config(font = ('Corbel', 12))
btn_buscar_archivo = tb.Button(insert_file_frame, text = 'Buscar archivo', command = buscarArchivo, style = 'info.TButton', cursor = 'hand2')
btn_archivo_borrar = tb.Button(insert_file_frame, text = 'Borrar archivo seleccionado', command = borrarArchivo, style = 'info.TButton', cursor = 'hand2')
lbl_seleccionar_archivo.pack(side = 'left')
btn_buscar_archivo.pack(side = 'left', padx = 10)
btn_archivo_borrar.pack(side = 'left')
insert_file_frame.pack(fill = 'x', padx = 20)
lbl_file = tb.Label(master = window, foreground = '#366da3')
lbl_file.config(font = ('Corbel', 10))
lbl_file.pack(fill = 'x', padx = 20, pady = 10)

# Game Stats / General Stats ComboBox
radiobutton_team_frame = tb.Frame(window)
opcion = StringVar()
rb_partido = tb.Radiobutton(radiobutton_team_frame, text = 'Partido', variable = opcion, value = 'Partido', cursor = 'hand2', style = 'info.Toolbutton', command = click_partido)
rb_Stats = tb.Radiobutton(radiobutton_team_frame, text = 'Estadisticas', variable = opcion, value = 'Estadisticas', cursor = 'hand2', style = 'info.Outline.Toolbutton', command = click_stats)
rb_partido.pack(side = 'left')
rb_Stats.pack(side = 'left', padx = 10)
rb_partido.invoke()
radiobutton_team_frame.pack(fill = 'x', padx = 23, pady = 10)

# Radiobuton select team
radiobutton_team_frame = tb.Frame(window)
opcion_equipo = StringVar()
rb_local = tb.Radiobutton(radiobutton_team_frame, text = 'Local', variable = opcion_equipo, value = 'Local', cursor = 'hand2', style = 'info.TRadiobutton')
rb_visitante = tb.Radiobutton(radiobutton_team_frame, text = 'Visitante', variable = opcion_equipo, value = 'Visitante', cursor = 'hand2', style = 'info.TRadiobutton')
rb_local.pack(side = 'left')
rb_visitante.pack(side = 'left', padx = 20)
rb_local.invoke()
radiobutton_team_frame.pack(fill = 'x', padx = 23, pady = 20)

# Confirm / Exit buttons
buttons_frame = tb.Frame(window)
btn_confirm = tb.Button(buttons_frame, text = 'Confirmar', command = confirmar, style = 'success.TButton', cursor = 'hand2').pack(side = 'right')
btn_exit = tb.Button(buttons_frame, text = 'Cerrar aplicación', command = salir, style = 'primary.TButton', cursor = 'hand2').pack(side = 'right', padx = 10)
buttons_frame.pack(pady = 20, padx = 30, fill = 'x')

window.mainloop()