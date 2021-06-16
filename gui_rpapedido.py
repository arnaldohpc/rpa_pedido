from PyQt5 import uic, QtWidgets
from os import getenv, path, remove, rename
import sys
import pandas as pd
import win32com.client

qtCreatorFile = "pedido.ui"  # Nombre del archivo aquí.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)


class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.progressBar.setValue(0)
        self.button2.setEnabled(False)
        self.button1.clicked.connect(self.get_xls)
        self.button2.clicked.connect(self.convert_file)

    def get_xls(self):
        # Se accede a la carpeta Onedrive del usuario \ pedidos.
        home = str(getenv("ONEDRIVE"))
        filePath, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open file", "{}\\Pedidos".format(home)
        )
        if filePath != "":
            self.data = pd.ExcelFile(str(filePath))
            self.comboBox.clear()
            self.comboBox.addItems(list(self.data.sheet_names))
            # Se guarda la dirección de la carpeta del Pedido.
            self.path = str(path.dirname(filePath))
            self.button2.setEnabled(True)
            self.fpath = str(filePath)

    def convert_file(self):
        # Separamos la Cabecera del Excel en un DataFrame
        dfcab = self.data.parse(
            sheet_name=self.comboBox.currentText(), index_col=None, header=None, nrows=8,
        )
        # Separamos el Detalle del Excel en un DataFrame,
        # seleccionamos solo las columnas de Articulo y
        # cantidad y eliminamos las celadas vacías.
        dfdet = (
            self.data.parse(
                sheet_name=self.comboBox.currentText(),
                usecols=["ARTÍCULO", "CANT. RED."],
                skiprows=9,
                dtype={"ARTÍCULO": str},
            )
            .dropna()
            .rename(columns={"ARTÍCULO": "Material", "CANT. RED.": "Qtd."})
        )
        # Llamamos a la función de carga de pedido y le
        # enviamos la cabecera y el detalle.
        self.load_order(dfcab, dfdet)

    def load_order(self, cab, det):
        # print(cab) # verificamos df cabecera
        # Carga de Pedido al SAP con SAP Gui Scripting.
        self.button2.setEnabled(False)
        # Cuardamos en variables los datos que utilizamos
        # en la Cabecera.
        # order_date = self.convert_date(cab.iloc[1, 2])
        client_id = str(cab.iloc[3, 2])
        delivery_date = self.convert_date(cab.iloc[2, 8])
        payment = str(cab.iloc[3, 8])
        # alternative_address = str(cab.iloc[4, 8])
        # uid = str(cab.iloc[5, 8])
        ###############################
        # verificar df de cabecera.
        # print(
        #     "{} | {} | {} | {} | {}".format(
        #         type(client_id),
        #         type(delivery_date),
        #         type(payment),
        #         type(alternative_address),
        #         type(uid),
        #     )
        # )
        # print(
        #     "{} | {} | {} | {} | {}".format(
        #         client_id, delivery_date, payment, alternative_address, uid
        #     )
        # )
        ############################################
        # Al detalles le pasamos la función para que
        # convierta a int las Cantidades.
        det = det.applymap(lambda x: int(x))
        # print(det) # verficar df detalle
        arch_xslx = r"template.xlsx"
        ruta_xlsx = self.path + r"\\" + arch_xslx
        det.to_excel(ruta_xlsx, sheet_name="template", index=False)

        # Bloque Try que ejecuta el script de Sap Gui.
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return

            connection = application.Children(0)
            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return
            # Carga la Cabecera.
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "{}".format("ZSD_EASY")
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxt[0]").text = "{}".format("TPY1")
            session.findById("wnd[0]/usr/ctxt[0]").caretPosition = 4
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxt[2]").text = "{}".format(client_id)
            self.progressBar.setValue(20)
            # session.findById("wnd[0]/usr/ctxt[1]"
            # ).text = "{sector}.format(sector="05")"
            # session.findById("wnd[0]/usr/ctxt[3]"
            # ).text = "{vendedor}.format(vendedor=uid)"
            session.findById("wnd[0]/usr/ctxt[3]").setFocus()
            session.findById("wnd[0]/usr/ctxt[3]").caretPosition = 8
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/rad[2]").setFocus()
            session.findById("wnd[0]/usr/rad[2]").select()
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[1]"
            ).text = "{}".format("ZVVE")
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[2]"
            ).text = "{}".format(delivery_date)
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[3]"
            ).text = "{}".format("CIF")
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[4]"
            ).text = "{}".format(payment)
            session.findById("wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[4]").setFocus
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB1/ssub/2/3/ctxt[4]"
            ).caretPosition = 4
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/tabsTAB/tabpTAB2").select()
            session.findById(
                "wnd[0]/usr/tabsTAB/tabpTAB2/ssub/2/5/cntlCC_ITEM/shellcont/shell"
            ).pressToolbarButton("IMP")
            self.progressBar.setValue(60)
            session.findById("wnd[1]/usr/ctxt[0]").text = "{}".format(self.path)
            session.findById("wnd[1]/usr/ctxt[1]").text = "{}".format(arch_xslx)
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/btn[0]").press()
            self.error = 0
        except:
            self.mensaje = str(sys.exc_info()[0])
            self.error = 1
        finally:
            session = None
            connection = None
            application = None
            SapGuiAuto = None

        self.show_dialog()
        self.button2.setEnabled(False)
        self.comboBox.clear()
        self.progressBar.setValue(0)
        remove(ruta_xlsx)

    def show_dialog(self):
        # Muestra un Mensaje emergente.
        if self.error == 1:
            QtWidgets.QMessageBox.about(self, "Error", self.mensaje)
            rename(self.fpath, str(self.fpath + ".err"))
        else:
            self.progressBar.setValue(100)
            QtWidgets.QMessageBox.about(self, "Terminado", "El archivo fue cargado.")
            rename(self.fpath, str(self.fpath + ".ok"))

    def convert_date(sefl, date):
        # Funciòn que convierte la fecha en formato dd.mm.yyyy
        rname = str(date.day) + "." + str(date.month) + "." + str(date.year)
        return rname


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
