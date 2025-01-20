from PyQt5.QtWidgets import QMainWindow, QApplication, QLineEdit, QHeaderView, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QDate, QRegExp
from PyQt5.QtGui import QRegExpValidator
from PyQt5 import uic

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, Frame, ListFlowable, ListItem
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen.canvas import Canvas

from num2words import num2words

import sys
import os
import re




class CustomLineEdit(QLineEdit):
    """ """

    def __init__(self):
        super().__init__()

        self.textEdited.connect(self.formatearMonto)


    def formatearMonto(self, texto):
        """Coloca puntos al monto cada tres dígitos."""

        print(texto)
        # dni = self.lineEdit_dni.text()

        # # Si no hay al menos 4 dígitos, no formateo
        # if not len(dni) > 3:
        #     return

        # # Formateo DNI con puntos cada tres dígitos
        # dniFormateado = ''
        # while len(dni) > 3:
        #     dniFormateado = '.' + dni[-3:] + dniFormateado
        #     dni = dni[:-3]
        # dniFormateado = dni + dniFormateado

        # # Actualizo texto en QLineEdit de dni
        # self.lineEdit_dni.setText(dniFormateado)




class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Cargo el archivo .ui
        uic.loadUi('ui/presupuestos.ui', self)

        # Inicializaciones varias
        self.setWindowTitle('Presupuestador') # Título de la ventana
        # self.setWindowIcon(QIcon('resources/icons/icon.ico')) # Icono de la ventana
        self.showMaximized() # Abro la ventana maximizada

        # Carpeta de guardado de presupuestos
        self.pdfFolder = 'D:\\Google Drive'

        # Configuro acciones disparadas por QPushButtons
        self.pushButton_adddetalle.clicked.connect(self.agregarDetalle)
        self.pushButton_deldetalle.clicked.connect(self.eliminarDetalle)
        self.pushButton_addmonto.clicked.connect(self.agregarMonto)
        self.pushButton_delmonto.clicked.connect(self.eliminarMonto)
        self.pushButton_vaciar.clicked.connect(self.vaciar)
        self.pushButton_guardar.clicked.connect(self.guardar)

        self.lineEdit_total.textChanged.connect(self.formatearMonto)

        self.tableWidget_detalles.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_montos.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)



        # Cargo ítems en el ComboBox de clientes
        self.cargarClientes()

        # Pongo fecha actual
        self.ponerFechaActual()


    def cargarClientes(self):
        # Creo la lista de clientes para agregar
        self.clientes = []
        with open('clientes.txt', 'r') as archivo:
            for linea in archivo:
                self.clientes.append(linea.strip())
            self.clientes.sort()

        self.comboBox_clientes.addItems(self.clientes) # Agrego todos los ítems de la lista

        # Establecer sin selección inicial
        self.comboBox_clientes.setCurrentIndex(-1)

        # print(f'"{self.comboBox_clientes.currentText()}"')


    def ponerFechaActual(self):
        self.dateEdit_fecha.setDate(QDate.currentDate())


    def agregarDetalle(self):
        """Agrega una fila al final de la tabla de detalles."""

        nroFilas = self.tableWidget_detalles.rowCount()
        self.tableWidget_detalles.insertRow(nroFilas)  # Inserta una fila al final


    def eliminarDetalle(self):
        """Elimina la fila seleccionada."""
        
        filaActual = self.tableWidget_detalles.currentRow() # Obtiene la fila seleccionada
        if filaActual != -1: # Me aseguro de que haya una fila seleccionada
            self.tableWidget_detalles.removeRow(filaActual)


    def agregarMonto(self):
        """Agrega una fila al final de la tabla de montos."""

        # Inserto una fila al final
        nroFilas = self.tableWidget_montos.rowCount()
        self.tableWidget_montos.insertRow(nroFilas)

        # Hago que aparezca el total si hay una fila
        if self.tableWidget_montos.rowCount() == 1:
            self.lineEdit_total.setPlaceholderText('$ 0')
            
        # Inserto el QLineEdit del monto en la  celda
        lineEdit = QLineEdit()
        lineEdit.setPlaceholderText('$ 0')
        lineEdit.setAlignment(Qt.AlignRight)
        self.tableWidget_montos.setCellWidget(nroFilas, 1, lineEdit)
        
        # Conecto señales al QLineEdit
        lineEdit.textEdited.connect(self.formatearMonto)
        lineEdit.textEdited.connect(self.actualizarTotal)


    def limpiarMonto(self, stringMonto):
        """Quita de un string de monto cualquier caracter que no sea un dígito."""
        
        return re.sub(
            r'\D',          # Regex (match donde el string NO contiene dígitos)
            '',             # Por lo que quiero reemplazar
            stringMonto     # Dónde buscar
        )


    def formatearMonto(self, texto):
        """Coloca puntos cada tres dígitos al monto."""

        sender = self.sender()

        # Quito cualquier caracter que no sea un dígito
        monto = self.limpiarMonto(sender.text())

        # Quito ceros al principio
        if len(monto):
            monto = str(int(monto))
        
        # Si no hay al menos 4 dígitos, no formateo
        if not len(monto) > 3:
            if not len(monto) or monto == '0':
                sender.clear()
            else:
                sender.setText(f'$ {monto}')
            return

        # Formateo el monto con puntos cada tres dígitos
        montoFormateado = ''
        while len(monto) > 3:
            montoFormateado = '.' + monto[-3:] + montoFormateado
            monto = monto[:-3]
        montoFormateado = monto + montoFormateado

        # Actualizo monto en QLineEdit
        sender.setText(f'$ {montoFormateado}')


    def actualizarTotal(self):
        """Realiza la sumatoria de los montos y actualiza Total."""

        # Al principio Total comienza vacio (string vacio)
        self.lineEdit_total.clear()

        # Recorro las filas y voy sumando montos
        for fila in range(self.tableWidget_montos.rowCount()):
            stringMonto = self.limpiarMonto(self.tableWidget_montos.cellWidget(fila, 1).text())

            if stringMonto: # Verifico que tenga un monto (porque si está el placeholder text hay un string vacio)
                monto = int(stringMonto)
                
                # Determino el valor actual de Total
                if self.lineEdit_total.text():
                    total = int(self.limpiarMonto(self.lineEdit_total.text()))
                else:
                    total = 0
            
                # Actualizo Total
                self.lineEdit_total.setText(f'$ {str(total + monto)}')


    def eliminarMonto(self):
        """Elimina la fila seleccionada de la tabla de montos."""

        filaActual = self.tableWidget_montos.currentRow() # Obtiene la fila seleccionada
        if filaActual != -1: # Me aseguro de que haya una fila seleccionada
            self.tableWidget_montos.removeRow(filaActual)

        # Quito el placeholder text del total si no hay filas de monto
        if not self.tableWidget_montos.rowCount():
            self.lineEdit_total.setPlaceholderText('')

        # Vuelvo a actualizar el total para quitar el monto eliminado
        self.actualizarTotal()


    def vaciar(self):
        self.cargarClientes()
        self.ponerFechaActual()
        self.lineEdit_titulo.clear()
        self.tableWidget_detalles.setRowCount(0)
        self.tableWidget_montos.setRowCount(0)
        self.lineEdit_total.clear()


    def guardar(self):
        """
        Agrega el cliente si no estaba en la lista, reúne los datos para armar el PDF,
        y genera el cuadro de diálogo para guardar el PDF.
        """

        # Verifico si el cliente actual existe
        clienteActual = self.comboBox_clientes.currentText()
        if clienteActual not in self.clientes:
            with open('clientes.txt', 'a') as archivo:
                archivo.write(clienteActual)

        # Obtengo una lista de los detalles
        listaDetalles = []
        for fila in range(self.tableWidget_detalles.rowCount()):
            item = self.tableWidget_detalles.item(fila, 0)
            listaDetalles.append(item.text() if item else '-')

        # Obtengo una lista de los montos
        listaMontos = []
        for fila in range(self.tableWidget_montos.rowCount()):
            item = self.tableWidget_montos.item(fila, 0)
            monto = self.tableWidget_montos.cellWidget(fila, 1).text()
            listaMontos.append((item.text() if item else '-', monto))

        # Armo un dicccionario para reunir la información que irá en el PDF
        datos = {
            'cliente': self.comboBox_clientes.currentText(),
            'fecha': self.dateEdit_fecha.date().toString('yyyy-MM-dd'),
            'titulo': self.lineEdit_titulo.text(),
            'iva': 'sin IVA' if self.radioButton_siniva.isChecked() else 'con IVA',
            'detalles': listaDetalles,
            'montos': listaMontos,
            'total': self.lineEdit_total.text()
        }

        # Defino el nombre del archivo
        defaultFilename = f"{datos['fecha']}_{datos['cliente']}_{datos['titulo']}.pdf"

        # Obtengo la carpeta por defecto que mostrará el QFileDialog de guardado
        if os.path.exists(self.pdfFolder): # Si existe la carpeta de Google Drive, selecciono esa
            defaultFolder = self.pdfFolder
        else:
            homeFolder = os.path.expanduser('~')
            for folder in ('Escritorio', 'Desktop'):
                desktopFolder = os.path.join(homeFolder, folder)
                if os.path.exists(desktopFolder):
                    defaultFolder = desktopFolder # Si no existe la carpeta de Google Drive, selecciono el Escritorio
            defaultFolder = homeFolder # Como alternativa final, uso la carpeta de usuario

        # Genero cuadro de diálogo para guardar el PDF
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getSaveFileName(
            self, 
            'Guardar PDF', 
            os.path.join(defaultFolder, defaultFilename), # Ubicación y nombre predeterminados
            'Archivos PDF (*.pdf)', 
            options=options
        )

        if filePath:
            try:
                # Genero el PDF en la ubicación seleccionada
                self.crearPdf(datos, filePath)
                
                # Muestro un mensaje de éxito
                self.mostrarMensajeGuardado(filePath)

                # Limpio todo
                self.vaciar()
            except Exception as e:
                # Muestro mensaje de error en caso de fallo
                QMessageBox.critical(self, 'Error', f'No se pudo generar el PDF:\n{str(e)}')


    def mostrarMensajeGuardado(self, filePath):
        """
        Muestra un mensaje tras el guardado del pdf, permitiendo abrir la carpeta de guardado.
        """

        # Creo el QMessageBox
        msgBox = QMessageBox(self)
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setWindowTitle('Éxito')
        msgBox.setText(f'El presupuesto se guardó correctamente en:')
        msgBox.setInformativeText(os.path.dirname(filePath)) # Muestro la carpeta contenedora, no el path completo del archivo
        
        # Añado botones de "Abrir carpeta" y "Cerrar"
        botonAbrir = msgBox.addButton('Abrir carpeta', QMessageBox.ActionRole)
        botonCerrar = msgBox.addButton('Cerrar', QMessageBox.RejectRole)

        # Configuro el cursor de mano para los botones
        botonAbrir.setCursor(Qt.PointingHandCursor)
        botonCerrar.setCursor(Qt.PointingHandCursor)

        # Muestro el mensaje
        msgBox.exec_()

        # Verifico si el usuario seleccionó "Abrir carpeta"
        if msgBox.clickedButton() == botonAbrir:
            os.startfile(os.path.dirname(filePath))


    def crearPdf(self, datos, filePath):
        """
        Construye el PDF del presupuesto y lo guarda en la ruta que le pasamos.
        """
        print(filePath)
        # Defino tamaños de fuente
        tamañoFuenteMontos = 11

        # Defino los márgenes de la página (en puntos)
        margenIzq, margenDer = 60, 60
        margenSup, margenInf = 40, 40

        # Configuración básica del documento
        doc = SimpleDocTemplate(filePath, 
                                pagesize=A4,
                                leftMargin=margenIzq,
                                rightMargin=margenDer,
                                topMargin=margenSup,
                                bottomMargin=margenInf)
        styles = getSampleStyleSheet()
        partesPdf = [] # Lista de elementos que se agregarán secuencialmente al PDF

        headingStyle = ParagraphStyle(
            name='EstiloTitulo',
            parent=styles['Heading2'],
            textColor='#0d0d0d', 
            spaceAfter=10
        )

        textStyle = ParagraphStyle(
            name='EstiloTexto',
            parent=styles['BodyText'],
            fontName='Helvetica',
            fontSize=11,
            textColor='#0d0d0d', 
            leading=15,  # Espaciado entre líneas
            spaceBefore=6,  # Espaciado antes del párrafo
            spaceAfter=6,  # Espaciado después del párrafo
            alignment=4  # 4 indica alineación justificada
        )

        bulletStyle = ParagraphStyle(
            'Bullet',
            parent=textStyle,
            bulletIndent=5,  # Espacio más pequeño para la viñeta
            leftIndent=-5,   # Reduce el espacio entre la viñeta y el texto
            fontSize=11,     # Tamaño de la fuente
            leading=14,      # Espaciado entre líneas
            alignment=4      # 4 indica alineación justificada

        )

        # partesPdf.append(Paragraph("Ancho disponible en linea (pt): 495.27559055118115 Ancho disponible en linea (pt): 495.27559055118115 Ancho disponible en linea (pt): 495.27559055118115", styles["BodyText"]))

        # ------------------------------------------ CLIENTE, FECHA, TITULO --------------------------------------
        
        # Cliente
        partesPdf.append(Paragraph(datos['cliente'], styles['Heading2']))
        
        # Título y fecha
        partesPdf.append(Paragraph(f"{datos['titulo']} ({datos['fecha']})", styles['Heading2']))

        # Añado espacio entre secciones
        partesPdf.append(Spacer(1, 25))

        # ---------------------------------------------- DETALLES ------------------------------------------------

        # Convierto los detalles en un ListFlowable con viñetas
        bulletList = ListFlowable(
            [ListItem(Paragraph(detalle, bulletStyle)) for detalle in datos['detalles']],
            bulletType='bullet' # Opción para viñetas (puedo usar "1" para numeración)
        )

        # Agrego la lista al documento
        partesPdf.append(bulletList)

        # Añado espacio entre secciones
        partesPdf.append(Spacer(1, 25))

        # ----------------------------------------------- MONTOS -------------------------------------------------

        # Calculo cuántos caracteres monoespaciados caben en una línea
        anchoDoc, _ = A4 # Tamaño de la página en puntos (solo me interesa el ancho)
        anchoDisponible = anchoDoc - margenIzq - margenDer
        canvas = Canvas(None) # Creo un canvas temporal en memoria para poder usar stringWidth(), que dá el ancho exacto
        anchoCaracter = canvas.stringWidth('A', fontName='Courier', fontSize=tamañoFuenteMontos) # Calculo el ancho de un carácter cualquiera
        nroCaracteres = int(anchoDisponible // anchoCaracter) # Número de caracteres monoespaciados que caben en una línea

        # Encuentro longitud máxima de monto
        longMaxMonto = len(datos['total']) # (Nunca será inferior que la del total)

        # Creo lista de filas para la tabla de montos
        tablaMontos = []
        for detalle, monto in datos['montos']:
            puntos = '.' * (nroCaracteres - len(detalle) - longMaxMonto - 4) # Calculo el relleno de puntos. El 4 fué prueba y error
            tablaMontos.append([detalle + ' ' + puntos, monto])

        # Agrego la fila del total a la lista de filas de montos 
        tablaMontos = tablaMontos + [[f"TOTAL ({datos['iva']})", datos['total']]]

        # Calculo ancho de columnas de la tabla
        anchoColumnaMonto = (longMaxMonto + 2) * anchoCaracter # Ancho máximo requerido (en puntos) por la columna "Monto"
        anchoColumnaDetalle = anchoDisponible - anchoColumnaMonto        

        # Creo la tabla y ajusto sus columnas
        tabla = Table(tablaMontos, colWidths=[anchoColumnaDetalle, anchoColumnaMonto])
        tabla.setStyle(TableStyle([ # Las tuplas son (columna, fila)
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Alineo los detalle de los montos a la izquierda
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'), # Alineo los montos a la derecha
            ('ALIGN', (0, -1), (0, -1), 'RIGHT'), # Alineo la celda del "TOTAL" a la derecha
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black), # Color del texto
            ('FONTNAME', (0, 0), (-1, -1), 'Courier'), # Tipo de fuente general
            ('FONTNAME', (0, -1), (-1, -1), 'Courier-Bold'), # Fuente negrita para el monto total
            ('FONTSIZE', (0, 0), (-1, -1), tamañoFuenteMontos), # Tamaño de fuente
            ('LINEABOVE', (-1, -1), (-1, -1), 1, colors.black) # Hago visible la parte superior de la celda inferior derecha
        ]))

        # Agrego la tabla al documento
        partesPdf.append(tabla)

        # Añado espacio entre secciones
        partesPdf.append(Spacer(1, 25))

        # -------------------------------------------- TOTAL A TEXTO ---------------------------------------------

        # Convierto el total a entero
        montoEntero = int(re.sub(r'\D', '', datos['total']))

        # Convierto el entero a texto
        montoComoTexto = num2words(montoEntero, lang='es').capitalize()

        partesPdf.append(Paragraph(f"<b>Son pesos:</b> {montoComoTexto} ({datos['iva']} incluido)", textStyle))


        # Genero el PDF
        doc.build(partesPdf)






# Initialize the app
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Establezco tema de aplicación (windowsvista, Windows, Fusion)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())