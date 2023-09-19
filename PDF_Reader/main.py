import PDF_Reader
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys
import PDF_Read_Func

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        QMainWindow.__init__(self)
        self.ui = PDF_Reader.Ui_MainWindow()
        self.ui.setupUi(self)

        # ============================================================================================================
        self.ui.PDF_Sel_Btn.clicked.connect(lambda : self.SelectFile(self.ui.PDF_Path_Txt))
        self.ui.Excel_Sel_Btn.clicked.connect(lambda : self.SelectFile(self.ui.Excel_Path_Txt))
        self.ui.StartBtn.clicked.connect(lambda : PDF_Read_Func.PDF_Read_Start(
            self.ui.PDF_Path_Txt.text(),
            self.ui.Excel_Path_Txt.text(),
            self.ui.progressBar,
            self.ui.PDF_Sel_Btn,
            self.ui.Excel_Sel_Btn,
            self.ui.StartBtn
        ))
        # ============================================================================================================

        self.show()

    def SelectFile(self, Path_Text):
        FileName = QFileDialog.getOpenFileName(self, 'Open file', './')[0]
        if FileName == '':
            Path_Text.setText('None')
        else:
            Path_Text.setText(FileName)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())