import sys, os , json
from PyQt6.QtWidgets import QApplication, QWidget,QVBoxLayout,QPushButton,QLineEdit,QLabel,QFormLayout,QComboBox,QSizePolicy
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QCloseEvent
from BasketTrade_ajout import Window
from datetime import datetime 
from Style import get_stylesheet

CONFIG_FILE = "config.json"


class  WindowUser(QWidget):
    user_folder = pyqtSignal(str,str)
    def __init__(self):
        super().__init__()
        self.can_close = False 
        self.setWindowTitle("Configuration")
        self.setFixedSize(400,300)
        self.setStyleSheet(get_stylesheet())

        layout = QVBoxLayout(self)
        users = ['user_a','user_b']
        folders = ['UAT','TEST','PROD']

        
        self.combobox_user = QComboBox()
        self.text_user = QLabel('Utilisateur : ')
        
        self.combobox_user.addItems(users)

        self.text_folder = QLabel('Dépôt : ')
        
        self.combobox_folder = QComboBox()
        self.combobox_folder.addItems(folders)

       

        self.button_valid = QPushButton('Valider')
        self.button_valid.clicked.connect(self.validation)

        layout.setSpacing(40)
        layout.setContentsMargins(8,8,8,8)
        layout.addWidget(self.text_user)
        layout.addWidget(self.combobox_user)
        layout.addWidget(self.text_folder)
        layout.addWidget(self.combobox_folder)
        layout.addWidget(self.button_valid)

    
    def validation (self):
        
        user = self.combobox_user.currentText()
        folder = self.combobox_folder.currentText()
        self.user_folder.emit(user,folder)
        self.can_close = True
        
        self.close()
    
    def closeEvent(self, event: QCloseEvent):
        if self.can_close == False :
            
            event.ignore()
        else : 
            
            event.accept()
    
class UseInfo():
    def __init__(self,user_id = '',folder_id = ''):
        self.user_id = user_id
        self.folder_id = folder_id
    
    def init(self,user,folder):
        self.user_id = user
        self.folder_id = folder

class character : 
    """Information sur les chemins de depots"""
    def __init__(self,use_info):
        super().__init__()

        with open(CONFIG_FILE,'r',encoding='utf-8') as f:
            dic = json.load(f)

        self.name_processed = dic['name_processed_folder']
        self.path_processed = dic['path_processed_file']
        self.file_processed = os.path.join(self.path_processed,self.name_processed)

        self.name_loggerfile = dic['name_loggerfile']
        self.path_logger = dic['path_logger']
        self.file_loggerpath = os.path.join(self.path_logger,self.name_loggerfile)

        
        #self.path_depot = dic['path_depot']
        #self.file_depotpath = os.path.join(self.path_depot,'')

        self.absolut_path = dic['absolut_path']

        self.user_id = use_info.user_id
        self.folder_id = use_info.folder_id

    
def main():
    app = QApplication(sys.argv)
    
    use_info = UseInfo('','')
    winuser = WindowUser()
    winuser.user_folder.connect(use_info.init)
    winuser.show()
    app.exec()

    char = character(use_info)
    window = Window(char)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

