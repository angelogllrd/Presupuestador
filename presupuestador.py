#######################################################################################################
#                                                                                                     #
#  ____   ____     ___   _____ __ __  ____   __ __    ___   _____ ______   ____  ___     ___   ____   #
# |    \ |    \   /  _] / ___/|  T  T|    \ |  T  T  /  _] / ___/|      T /    T|   \   /   \ |    \  #
# |  o  )|  D  ) /  [_ (   \_ |  |  ||  o  )|  |  | /  [_ (   \_ |      |Y  o  ||    \ Y     Y|  D  ) #
# |   _/ |    / Y    _] \__  T|  |  ||   _/ |  |  |Y    _] \__  Tl_j  l_j|     ||  D  Y|  O  ||    /  #
# |  |   |    \ |   [_  /  \ ||  :  ||  |   |  :  ||   [_  /  \ |  |  |  |  _  ||     ||     ||    \  #
# |  |   |  .  Y|     T \    |l     ||  |   l     ||     T \    |  |  |  |  |  ||     |l     !|  .  Y #
# l__j   l__j\_jl_____j  \___j \__,_jl__j    \__,_jl_____j  \___j  l__j  l__j__jl_____j \___/ l__j\_j #
#                                                                                                     #
#      Aplicación sencilla y a medida para generar PDFs de presupuestos con un formato estándar.      #
#                                                                                                     #
# Autor: Angelo Gallardi (angelogallardi@gmail.com)                                                   #
# Año: 2025                                                                                           #
# Versión: 1.0                                                                                        #
#                                                                                                     #
#######################################################################################################


from PyQt5.QtWidgets import QMainWindow, QApplication, QLineEdit, QHeaderView, QFileDialog, QDialog, QMessageBox
from PyQt5.QtCore import Qt, QDate, QSettings
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import uic

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, ListFlowable, ListItem
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen.canvas import Canvas

from num2words import num2words

import sys
import os
import re


class ConfiguracionDialog(QDialog):
    def __init__(self, defaultFolderPath, settings):
        super().__init__()

        # Cargo el archivo .ui
        uic.loadUi('ui/settings.ui', self)

        # Inicializaciones varias
        self.setWindowIcon(QIcon('resources/icons/icon.ico')) # Icono de la ventana
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint) # Quito signo "?" de barra de título
        self.pushButton_seleccionar.clicked.connect(self.seleccionarCarpeta)
        self.pushButton_cerrar.clicked.connect(self.close) # Uso close() para cerrar la ventana

        # Guardo los parámetros recibidos
        self.defaultFolderPath = defaultFolderPath
        self.settings = settings

        # Muestro la ruta actual en el line edit
        self.lineEdit_carpeta.setText(self.settings.value('pdfs_folder', self.defaultFolderPath))


    def seleccionarCarpeta(self):
        """Permite al usuario seleccionar una carpeta y la guarda en QSettings."""

        folder = QFileDialog.getExistingDirectory(self, 'Seleccionar Carpeta', self.settings.value('pdfs_folder', self.defaultFolderPath))
        if folder:
            self.settings.setValue('pdfs_folder', folder) # Actualizo la ruta en QSettings
            self.lineEdit_carpeta.setText(folder) # La muestro en el line edit



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Cargo el archivo .ui
        uic.loadUi('ui/presupuestador.ui', self)

        # Inicializaciones varias
        self.setWindowTitle('Presupuestador') # Título de la ventana
        self.setWindowIcon(QIcon('resources/icons/icon.ico')) # Icono de la ventana
        self.showMaximized() # Abro la ventana maximizada

        # Variables varias
        self.anchoColumnaMonto = 160

        # Configuro acciones disparadas por QPushButton
        self.pushButton_adddetalle.clicked.connect(self.agregarDetalle)
        self.pushButton_deldetalle.clicked.connect(self.eliminarDetalle)
        self.pushButton_addmonto.clicked.connect(self.agregarMonto)
        self.pushButton_delmonto.clicked.connect(self.eliminarMonto)
        self.pushButton_vaciar.clicked.connect(self.vaciar)
        self.pushButton_guardar.clicked.connect(self.guardar)
        self.pushButton_config.clicked.connect(self.abrirConfiguracion)

        # Configuro acciones disparadas por QLineEdit
        self.lineEdit_total.textChanged.connect(self.formatearMonto)

        # Configuro acciones disparadas por QTableWidget
        self.tableWidget_detalles.cellChanged.connect(self.celdaCambiada) # Señal que se emite cada vez que cambia el contenido de una celda
        self.tableWidget_montos.cellChanged.connect(self.celdaCambiada)

        # Trato la ruta de guardado de los presupuestos
        self.settings = QSettings('COMET', 'Presupuestador') # Creo QSettings para guardar/recuperar configuración (se almacena en registro de Windows)
        self.defaultFolderPath = self.obtenerPathEscritorio() # Determino carpeta correcta del escritorio y la tomo como ruta de guardado por defecto
        self.pdfsFolderPath = self.settings.value('pdfs_folder', self.defaultFolderPath) # Recupero carpeta guardada previamente, o uso carpeta por defecto

        # Verifico si tengo el archivo de clientes, y sino lo creo
        self.clientesFilePath = self.verificarArchivoClientes()

        # Cargo ítems en el ComboBox de clientes
        self.cargarClientes()

        # Pongo fecha actual
        self.ponerFechaActual()

        # Ajusto el ancho de las columnas de ambas tablas equitativamente
        self.formatearTablas()


    def celdaCambiada(self, row, column):
        """Ajusta el alto de las filas al contenido para que todo el texto sea visible."""

        # Detecto cuál table widget lo llamó
        sender = self.sender()

        # Ajusto el alto de las filas al contenido
        sender.resizeRowsToContents()

        # Evito que resizeRowsToContents() me achique la fila más que el alto original
        for fila in range(sender.rowCount()):
            if sender.rowHeight(fila) < 30: # 30 es el alto original
                sender.setRowHeight(fila, 30)


    def verificarArchivoClientes(self):
        """
        Verifica si ya existe clientes.txt, o lo crea. El directorio "Presupuestador" donde
        se coloca clientes.txt debe estar en C:/Users/%USERNAME%/AppData/Roaming para que
        permita modificar el .txt (en Archivos de programa no está permitido). Ver, por ej,
        https://getgreenshot.org/faq/where-does-greenshot-store-its-configuration-settings/
        que hace lo mismo.
        """

        # Obtengo la ruta a AppData/Roaming
        appdataFolderPath = os.getenv('APPDATA')

        # Ruta a la carpeta de clientes
        clientesFolderPath = os.path.join(appdataFolderPath, 'Presupuestador')
        os.makedirs(clientesFolderPath, exist_ok=True) # Creo la carpeta si no existe

        # Ruta completa al archivo de clientes
        clientesFilePath = os.path.join(clientesFolderPath, 'clientes.txt')

        # Verifico si clientes.txt ya existe
        if not os.path.exists(clientesFilePath): # Verifico si el archivo no existe
            with open(clientesFilePath, 'w') as f: # Creo el archivo vacío
                f.write('')

        return clientesFilePath


    def cargarClientes(self):
        """Carga los clientes desde el .txt y los pone en el combo box."""

        # Limpio el contenido actual del combo box
        self.comboBox_clientes.clear()

        # Creo la lista de clientes para agregar
        self.clientes = []
        with open(self.clientesFilePath, 'r') as archivo:
            for linea in archivo:
                self.clientes.append(linea.strip()) # Elimino saltos de línea
        
        # Ordeno los clientes
        self.clientes.sort()

        # Agrego todos los ítems de la lista
        self.comboBox_clientes.addItems(self.clientes)

        # Establezco que no haya uno seleccionado
        self.comboBox_clientes.setCurrentIndex(-1)


    def ponerFechaActual(self):
        """Pone la fecha actual en el QDateEdit."""

        self.dateEdit_fecha.setDate(QDate.currentDate())


    def formatearTablas(self):
        """Ajusta el ancho de las columnas y el alto de los encabezados de ambas tablas."""

        self.tableWidget_detalles.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # La única columna ocupa el espacio disponible
        self.tableWidget_montos.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch) # La 1ra columna ocupa el espacio disponible que le deja la 2da
        self.tableWidget_montos.setColumnWidth(1, self.anchoColumnaMonto) # Fijo el ancho de la segunda columna al mismo del QLineEdit del monto

        # Modifico la altura de los encabezados horizontales de ambas tablas
        self.tableWidget_detalles.horizontalHeader().setFixedHeight(23)
        self.tableWidget_montos.horizontalHeader().setFixedHeight(23)


    def agregarDetalle(self):
        """Agrega una fila al final de la tabla de detalles."""

        # Inserto una fila al final
        nroFilas = self.tableWidget_detalles.rowCount()
        self.tableWidget_detalles.insertRow(nroFilas)


    def eliminarDetalle(self):
        """Elimina la fila seleccionada de la tabla de detalles."""

        filaActual = self.tableWidget_detalles.currentRow() # Obtengo la fila seleccionada
        if filaActual != -1: # Me aseguro de que haya una fila seleccionada
            self.tableWidget_detalles.removeRow(filaActual)


    def agregarMonto(self):
        """Agrega una fila al final de la tabla de montos."""

        # Inserto una fila al final
        nroFilas = self.tableWidget_montos.rowCount()
        self.tableWidget_montos.insertRow(nroFilas)

        # Hago que aparezca el placeholder text del total si hay una fila
        if self.tableWidget_montos.rowCount() == 1:
            self.lineEdit_total.setPlaceholderText('$ 0')

        # Creo un QLineEdit y le doy formato
        lineEdit = QLineEdit()
        lineEdit.setPlaceholderText('$ 0')
        lineEdit.setAlignment(Qt.AlignRight)
        lineEdit.setFixedWidth(self.anchoColumnaMonto)

        # Configuro la fuente del monto
        font = QFont('Courier', 11)
        font.setBold(True)
        lineEdit.setFont(font)

        # Inserto el QLineEdit en la celda del monto
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

        # Obtengo el line edit que llamó al método
        sender = self.sender()

        # Quito cualquier caracter que no sea un dígito
        monto = self.limpiarMonto(sender.text())

        # Quito posibles ceros al principio
        if len(monto):
            monto = str(int(monto))
        
        # Si no hay al menos 4 dígitos, no formateo
        if not len(monto) > 3:
            if not len(monto) or monto == '0': # Hago clear() para mostrar el placeholder text cuando: borré y quedó sin monto (para evitar que
                sender.clear()                 # quede "$" solo) y cuando pongo un 0 (no suma al total y ya está en el placeholder text)
            else:
                sender.setText(f'$ {monto}') # Si tengo un monto válido con 1 a 3 dígitos, le pongo el "$""
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

        filaActual = self.tableWidget_montos.currentRow()
        if filaActual != -1: # Me aseguro de que haya una fila seleccionada
            self.tableWidget_montos.removeRow(filaActual)

        # Quito el placeholder text del total si no hay filas de monto
        if not self.tableWidget_montos.rowCount():
            self.lineEdit_total.setPlaceholderText('')

        # Vuelvo a actualizar el total para quitar el monto eliminado
        self.actualizarTotal()


    def vaciar(self):
        """Deja todo como al abrir el programa."""

        self.cargarClientes()
        self.ponerFechaActual()
        self.lineEdit_titulo.clear()
        self.tableWidget_detalles.setRowCount(0)
        self.tableWidget_montos.setRowCount(0)
        self.lineEdit_total.clear()
        self.lineEdit_total.setPlaceholderText('')
        self.deseleccionarRadioButtons()


    def deseleccionarRadioButtons(self):
        """
        Desactiva y restaura la eutoexclusividad de los radio buttons para poder
        deseleccionarlos.
        """

        self.radioButton_siniva.setAutoExclusive(False)
        self.radioButton_coniva.setAutoExclusive(False)

        self.radioButton_siniva.setChecked(False)
        self.radioButton_coniva.setChecked(False)

        self.radioButton_siniva.setAutoExclusive(True)
        self.radioButton_coniva.setAutoExclusive(True)


    def guardar(self):
        """
        * Verifica si se completó lo necesario y de lo contrario lo informa.
        * Agrega el cliente presente en el combo box si no estaba en la lista.
        * Reúne los datos para armar el PDF.
        * Genera el cuadro de diálogo para guardar el PDF.
        """

        # Verifico que se haya completado lo necesario
        mensaje = ''
        hayProblemas = False
        cantProblemas = 0
        if not self.comboBox_clientes.currentText().strip():
            mensaje += '- Falta <b>Cliente</b><br>'
            hayProblemas = True
            cantProblemas += 1
        if not self.lineEdit_titulo.text().strip():
            mensaje += '- Falta <b>Título del trabajo</b><br>'
            hayProblemas = True
            cantProblemas += 1
        if not self.radioButton_siniva.isChecked() and not self.radioButton_coniva.isChecked():
            mensaje += '- Falta seleccionar <b>IVA</b><br>'
            hayProblemas = True
            cantProblemas += 1
        if not self.lineEdit_total.text():
            mensaje += '- No hay montos en la tabla <b>Montos</b><br>'
            hayProblemas = True
            cantProblemas += 1

        # Muestro el QMessageBox si hay problemas
        if hayProblemas:
            mensaje = f"PROBLEMA{'S' if cantProblemas > 1 else ''}:<br>" + mensaje
            self.mostrarAdvertencia('Advertencia', mensaje)
            return

        # Verifico si el cliente actual existe, y sino lo agrego a clientes.txt
        clienteActual = self.comboBox_clientes.currentText()
        if clienteActual not in self.clientes:
            with open(self.clientesFilePath, 'a') as archivo:
                archivo.write(clienteActual + '\n')

        # Obtengo una lista de los detalles
        listaDetalles = []
        for fila in range(self.tableWidget_detalles.rowCount()):
            item = self.tableWidget_detalles.item(fila, 0)
            if item and item.text().strip():
                listaDetalles.append(item.text()[0].upper() + item.text()[1:])

        # Obtengo una lista de los montos
        listaMontos = []
        for fila in range(self.tableWidget_montos.rowCount()):
            item = self.tableWidget_montos.item(fila, 0)
            monto = self.tableWidget_montos.cellWidget(fila, 1).text()
            if monto: # Solo considero válida la fila si hay un monto
                listaMontos.append((item.text()[0].upper() + item.text()[1:] if item and item.text().strip() else '-', monto))

        # Armo un dicccionario con la información que irá en el PDF
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

        # Genero cuadro de diálogo para guardar el PDF
        options = QFileDialog.Options()
        pdfFilePath, _ = QFileDialog.getSaveFileName(
            self, 
            'Guardar PDF', 
            os.path.join(self.pdfsFolderPath, defaultFilename), # Ubicación y nombre predeterminados
            'Archivos PDF (*.pdf)', 
            options=options
        )

        if pdfFilePath:
            try:
                # Genero el PDF en la ubicación seleccionada
                self.crearPdf(datos, pdfFilePath)

                # Muestro un mensaje de éxito
                self.mostrarMensajeGuardado(pdfFilePath)

                # Limpio todo
                self.vaciar()
            except Exception as e:
                # Muestro mensaje de error en caso de fallo
                QMessageBox.critical(self, 'Error', f'No se pudo generar el PDF:\n{str(e)}')


    def abrirConfiguracion(self):
        """Crea una instancia del QDialog de "Configuración" y lo muestra."""

        dialog = ConfiguracionDialog(self.defaultFolderPath, self.settings)
        dialog.exec_() # Muestra el diálogo en modo modal

        # Actualizo la ruta de guardado después de que se cierre el diálogo
        self.pdfsFolderPath = self.settings.value('pdfs_folder', self.pdfsFolderPath)


    def obtenerPathEscritorio(self):
        """Determina el nombre correcto de la carpeta de escritorio y devuelve su ruta."""

        homeFolderPath = os.path.expanduser('~')
        for folder in ('Escritorio', 'Desktop'):
            desktopFolderPath = os.path.join(homeFolderPath, folder)
            if os.path.exists(desktopFolderPath):
                return desktopFolderPath
        return homeFolderPath # Como alternativa, uso la carpeta de usuario


    def mostrarAdvertencia(self, titulo, mensaje):
        """Muestra un cuadro de diálogo de advertencia con el título y mensaje especificados."""

        msgBox = QMessageBox(self)
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setWindowTitle(titulo)
        msgBox.setTextFormat(Qt.RichText) # Permite interpretar HTML (para que sea visible la negrita)
        msgBox.setText(mensaje)
        msgBox.setStandardButtons(QMessageBox.Ok)

        # Accedo al botón OK y establezco el cursor de mano
        botonOk = msgBox.button(QMessageBox.Ok)
        botonOk.setCursor(Qt.PointingHandCursor)

        # Muestro el mensaje
        msgBox.exec_()


    def mostrarMensajeGuardado(self, pdfFilePath):
        """
        Muestra un mensaje tras el guardado del PDF, permitiendo abrir la carpeta de guardado,
        o abrir el PDF.
        """

        # Creo el QMessageBox
        msgBox = QMessageBox(self)
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setWindowTitle('Éxito')
        msgBox.setText(f'El presupuesto se guardó correctamente en:')
        msgBox.setInformativeText(os.path.dirname(pdfFilePath)) # Muestro la carpeta contenedora, no el path completo del archivo

        # Añado botones de "Abrir carpeta", "Abrir PDF" y "Cerrar"
        botonAbrirCarpeta = msgBox.addButton('Abrir carpeta', QMessageBox.ActionRole)
        botonAbrirArchivo = msgBox.addButton('Abrir PDF', QMessageBox.ActionRole)
        botonCerrar = msgBox.addButton('Cerrar', QMessageBox.RejectRole)

        # Configuro el cursor de mano para los botones
        botonAbrirCarpeta.setCursor(Qt.PointingHandCursor)
        botonAbrirArchivo.setCursor(Qt.PointingHandCursor)
        botonCerrar.setCursor(Qt.PointingHandCursor)

        # Muestro el mensaje
        msgBox.exec_()

        # Verifico si el usuario seleccionó "Abrir carpeta" o "Abrir PDF"
        if msgBox.clickedButton() == botonAbrirCarpeta:
            os.startfile(os.path.dirname(pdfFilePath))
        elif msgBox.clickedButton() == botonAbrirArchivo:
            os.startfile(pdfFilePath)


    def crearPdf(self, datos, pdfFilePath):
        """Construye el PDF del presupuesto y lo guarda en la ruta que le pasamos."""

        # Defino tamaños de fuente
        tamañoFuenteMontos = 10

        # Defino los márgenes de la página (en puntos)
        margenIzq, margenDer = 60, 60
        margenSup, margenInf = 40, 40

        # Configuración básica del documento
        doc = SimpleDocTemplate(pdfFilePath, 
                                pagesize=A4,
                                leftMargin=margenIzq,
                                rightMargin=margenDer,
                                topMargin=margenSup,
                                bottomMargin=margenInf)
        styles = getSampleStyleSheet()
        partesPdf = [] # Lista de elementos que se agregarán secuencialmente al PDF

        introStyle = ParagraphStyle(
            name='EstiloIntro',
            parent=styles['Heading2'],
            fontName='Helvetica',
            textColor='#1a1a1a',
            spaceBefore=6,
        )

        textStyle = ParagraphStyle(
            name='EstiloTexto',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=11,
            textColor='#1a1a1a', 
            leading=14,  # Espaciado entre líneas
            spaceBefore=6,  # Espaciado antes del párrafo
            spaceAfter=6,  # Espaciado después del párrafo
            alignment=4,  # 4 indica alineación justificada
        )

        bulletStyle = ParagraphStyle(
            'Bullet',
            parent=textStyle,
            bulletIndent=5,  # Espacio más pequeño para la viñeta
            leftIndent=-8,   # Reduce el espacio entre la viñeta y el texto
        )

        estiloSecciones = ParagraphStyle(
            name='EstiloSecciones',
            parent=styles['Heading2'],
            backColor=colors.gainsboro,
            textColor='#1a1a1a',
            # borderWidth=0.5,
            # borderColor='#1a1a1a',
            borderPadding=2,
            spaceAfter=12 # Espacio después del párrafo
        )

        # ------------------------------------------------- LOGO -------------------------------------------------

        # Cargo el logo del taller
        rutaLogo = 'resources/img/logo.png'
        logo = Image(rutaLogo)
        logo.hAlign = 'RIGHT' # Alineo el logo a la derecha

        # Obtengo el ancho y alto original de la imagen
        anchoOriginal, altoOriginal = logo.imageWidth, logo.imageHeight

        # Defino un ancho deseado
        anchoDeseado = 80

        # Calculo el alto proporcional al ancho deseado para respetar el aspect ratio
        aspectRatio = altoOriginal / anchoOriginal
        altoDeseado = anchoDeseado * aspectRatio

        # Aplico las dimensiones ajustadas al logo
        logo.drawWidth = anchoDeseado
        logo.drawHeight = altoDeseado

        # Añado espacio negativo para elevar el logo por encima del márgen superior
        partesPdf.append(Spacer(1, -24))

        # Agrego el logo al contenido del PDF
        partesPdf.append(logo)

        # ------------------------------------------ CLIENTE, FECHA, TITULO --------------------------------------
        
        # Cliente
        partesPdf.append(Paragraph(f"<b>Cliente:</b> {datos['cliente']}", introStyle))
        
        # Título
        partesPdf.append(Paragraph(f"<b>Presupuesto para:</b> {datos['titulo']}", introStyle))

        # Fecha
        fechaQdate = QDate.fromString(datos['fecha'], 'yyyy-MM-dd') # Convierto el string a un objeto QDate
        fechaConvertida = fechaQdate.toString('dd-MM-yyyy')
        partesPdf.append(Paragraph(f"<b>Fecha:</b> {fechaConvertida}", introStyle))

        # Añado espacio entre secciones
        partesPdf.append(Spacer(1, 25))

        # ---------------------------------------------- DETALLES ------------------------------------------------

        if datos['detalles']: 
            partesPdf.append(Paragraph('Detalles del trabajo', estiloSecciones))

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

        partesPdf.append(Paragraph('Presupuesto', estiloSecciones))

        # Calculo cuántos caracteres monoespaciados caben en un renglón
        anchoDoc, _ = A4 # Tamaño de la página en puntos (solo me interesa el ancho)
        anchoDisponible = anchoDoc - margenIzq - margenDer # Espacio para texto de un renglón en puntos
        canvas = Canvas(None) # Creo un canvas temporal en memoria para poder usar stringWidth(), que dá el ancho exacto
        anchoCaracter = canvas.stringWidth('A', fontName='Courier', fontSize=tamañoFuenteMontos) # Calculo el ancho de un caracter monoespaciado cualquiera
        maxCaracteresRenglon = int(anchoDisponible // anchoCaracter) # Número de caracteres monoespaciados que caben en un renglón

        # Encuentro longitud máxima de monto
        maxCaracteresMonto = len(datos['total']) # (Siempre será menor o igual que la del total)

        # Encuentro espacio disponible para el concepto
        maxCaracteresConcepto = maxCaracteresRenglon - maxCaracteresMonto - 4 # Resto 4 porque: 2 caracteres de concepto se solapaban con el monto (maxCaracteresRenglon
                                                                              # resulta en 79, pero solo entran 77 caracteres en el renglón), y otros 2 que dan el espacio
                                                                              # entre el concepto y el signo "$" del monto.

        # Creo lista de filas para la tabla de montos
        tablaMontos = []
        for concepto, monto in datos['montos']:
            palabras = concepto.split()
            renglonesConcepto = ''
            ultRenglonConcepto = ''

            # Formo el string del concepto ajustado al ancho de la columna (dividido en líneas)
            for palabra in palabras:
                if len(ultRenglonConcepto + palabra) <= maxCaracteresConcepto: # La palabra entra en el renglón actual
                    ultRenglonConcepto += palabra + ' '
                else: # La palabra no entra, la pongo en un nuevo renglón
                    renglonesConcepto += ultRenglonConcepto.rstrip() + '\n' # Inserto el nuevo renglón a los anteriores
                    ultRenglonConcepto = palabra + ' ' # Reinicio el último renglón
            ultRenglonConcepto = ultRenglonConcepto.rstrip() # Quito el último espacio después de la última palabra del último renglón
            renglonesConcepto += ultRenglonConcepto # Inserto el último renglón a los anteriores

            # Calculo cuántos puntos debo agregar al último renglón
            puntos = '.' * (maxCaracteresRenglon - len(ultRenglonConcepto) - maxCaracteresMonto - 4) # Resto 4 por lo mismo de más arriba en "maxCaracteresConcepto"

            tablaMontos.append([renglonesConcepto + puntos, monto])

        # Agrego la fila del total a la lista de filas de montos 
        tablaMontos = tablaMontos + [[f"TOTAL ({datos['iva']})", datos['total']]]

        # Calculo ancho de columnas de la tabla
        anchoColumnaMonto = (maxCaracteresMonto + 2) * anchoCaracter # Ancho máximo requerido (en puntos) por la columna "Monto"
        anchoColumnaDetalle = anchoDisponible - anchoColumnaMonto        

        # Creo la tabla y ajusto sus columnas
        tabla = Table(tablaMontos, colWidths=[anchoColumnaDetalle, anchoColumnaMonto])
        tabla.setStyle(TableStyle([ # Las tuplas son (columna, fila)
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Alineo los conceptos de los montos a la izquierda
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
        partesPdf.append(Spacer(1, 10))

        # -------------------------------------------- TOTAL COMO TEXTO ---------------------------------------------

        # Convierto el total a entero
        montoEntero = int(re.sub(r'\D', '', datos['total']))

        # Convierto el entero a texto
        montoComoTexto = num2words(montoEntero, lang='es').capitalize()

        partesPdf.append(Paragraph(f"<b>Son pesos:</b> {montoComoTexto} ({datos['iva']} incluido).", textStyle))


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