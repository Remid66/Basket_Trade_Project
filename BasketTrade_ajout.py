import shutil, os, subprocess,sys

from PyQt6.QtWidgets import QWidget,  QHBoxLayout, QVBoxLayout, QPushButton, QLabel, QLineEdit,QListWidget, QMessageBox
from PyQt6.QtCore import Qt , pyqtSignal, QThread
from PyQt6.QtGui import QDropEvent
import pymssql

from datetime import datetime

# =================================================================================================
# =                                       SECTION: THREAD
# =================================================================================================

class Worker(QThread):
    def __init__(self,info,parent = None):
        super().__init__(parent)
        self.info = info
        
    
    def run(self):
        script_path = os.path.abspath("Verification.py")
        try :
            result = subprocess.run([sys.executable, script_path,*self.info],
                                    stdout= subprocess.PIPE,
                                    stderr = subprocess.PIPE,
                                    text = True)
            print("STDOUT :", result.stdout)
            print("STDERR :",result.stderr)
        except Exception as e:
            print("Erreur lors du lancement de Verification.py")


# =================================================================================================
# =                                       SECTION: FENETRE
# =================================================================================================

class Window(QWidget):
    def __init__(self,character):
        super().__init__()
        self.setWindowTitle("Basket Trade")
        self.resize(500,300)
        self.character = character
        

        # -- Widgets --
        self.searchbox = SearchBox()

        self.searchbox.results_ready.connect(self.update_list)
        self.drop_area = DropLabel(character,"Deposit File")
        self.clear_button = ClearButton()
        self.accept_button = AcceptButton()

        self.list_widget = QListWidget()
        self.list_widget_value = None
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        
        self.clear_button.clicked.connect(self.drop_area.clearFile)

        self.accept_button.clicked.connect(self.verifacceptbuttonn)
    
        # -- Layout --
        
        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.clear_button)
        bottom_layout.addWidget(self.accept_button)

        # Layout principal
        main_layout= QVBoxLayout()
        main_layout.addWidget(self.searchbox)
        main_layout.addWidget(self.list_widget)
        
        
        main_layout.addWidget(self.drop_area)
        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

    def verifacceptbuttonn(self):
        if self.list_widget_value:
           
            self.drop_area.sendtoverif(self.list_widget_value)
            self.open_verif()
        else: 
            QMessageBox.warning(self, 'Error', 'Error : missing element')

    def on_item_clicked(self,item):
        self.list_widget_value = item.text()

    def update_list(self,items):
        self.list_widget.clear()
        self.list_widget.addItems(items)

    def open_verif(self):
            if self.drop_area.file_path is not None:
                self.verif_window = VerifWindow(self.drop_area.data, self,self.character)
                self.verif_window.show()
            else:
                QMessageBox.warning(self, 'Error', 'Error : File missing')
    
    def closeEvent(self, event):
        if self.searchbox.conn:
            self.searchbox.conn.close()
            print("Connexion SQL fermée")
        super().closeEvent(event)

class VerifWindow(QWidget):
    def __init__(self, data ,Window, character, parent = None):

        super().__init__(parent)
        self.setWindowTitle("Verification")
        self.setGeometry(100,100,300,200)

        self.main_window = Window
        self.data = data
        self.character = character

        self.accept_button = AcceptButtonVerif()
        self.cancel_button = CancelButtonVerif()
                  
        self.cancel_button.clicked.connect(self.cancel)
        self.accept_button.clicked.connect(self.sendFile)
            
        layout = QVBoxLayout()
        self.label = QLabel(f"Data : <b>{self.data.name}</b> <br> File : <b>{self.data.file_name}</b> <br> New File : <b>{self.data.subfolder_name}</b>")

        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.cancel_button)
        bottom_layout.addWidget(self.accept_button)


        layout.addWidget(self.label)
        layout.addLayout(bottom_layout)


        self.setLayout(layout) 

    def cancel (self):
            self.main_window.drop_area.clearFile()
            self.close() 

    def sendFile(self):

        
        os.makedirs(self.data.subfolder_path, exist_ok=True)

        new_file_path = shutil.copy(self.data.file_name, self.data.target)
        
        self.info = [new_file_path,self.data.name,self.data.time,self.data.subfolder_path,self.data.file_loggerpath,self.data.user_id,self.data.folder_id,self.data.absolut_path]

        self.worker = Worker(self.info)
        self.worker.start()


        self.main_window.drop_area.clearFile()
        self.close()
        
# =================================================================================================
# =                                       SECTION: WIDGET
# =================================================================================================

class SearchBox(QLineEdit):
    results_ready = pyqtSignal(list)

    def __init__(self, parent = None):
        super().__init__(parent)

        self.textChanged.connect(self.on_text_changed)

        self.conn = pymssql.connect(
                    server = '',
                    user = '',
                    password= '',
                    database=''
                )
        self.cursor = self.conn.cursor()

        self.setPlaceholderText("Research Client :")

    def on_text_changed(self,text):
        if text.strip():
            results  = self.search_in_db(text)
            self.results_ready.emit(results)
        else:
            self.results_ready.emit([])
        
    def search_in_db(self,keyword):
        """Cherche le mot clé"""
        if not self.cursor:
            return []
        try:
            query = f"""
                SELECT [Description] 
                FROM TBLCLIENTS 
                WHERE [Description] LIKE %s
                OR [Ul_id] LIKE %s
                OR [ExternalIds] LIKE %s
                AND [Active]  != 'false'
                """

            param = ('%' + keyword + '%', '%' + keyword + '%', '%' + keyword + '%' )
            self.cursor.execute(query,param)
            rows = self.cursor.fetchall()
            return [row[0] for row in rows]
        except Exception as e:
            print ("Erreur SQL :", e)
            return []  

class DropLabel(QLabel):
    fileDropped = pyqtSignal(str)

    def __init__(self, character , text = "Deposit File", parent = None):
        super().__init__(text)

        self.default_text = text
        self.data = data()
        self.character = character

        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(200,100)
        self.setAutoFillBackground(True)
        self.setStyleSheet("""
                           QLabel { border: 2px dashed gray;
                                    border-radius: 10px;
                                    background-color: #f0f0f0;
                                    font-size: 16px; 
                                    padding: 20px;
                           }
                           QLabel: hover{
                                    border: 2px dashed #0078d7;}
                           """)
        self.setAcceptDrops(True)
        self.file_path = None
    
    def dragEnterEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            self.setStyleSheet("""QLabel { border: 2px dashed #3399ff ;
                                       background-color: #e6f2ff;
                                       color: #333333;
                                       font-size: 14px; 
                                       padding: 20px;}""")
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""QLabel { border: 2px dashed #aaaaaa;
                                       background-color: #f9f9f9;
                                       color: #555555;
                                       font-size: 14px; 
                                       padding: 20px;}""")
           
    def dropEvent(self, event:QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                self.file_path = urls[0].toLocalFile()
                self.setText(f"Fichier déposé:\n {self.file_path}")
                self.fileDropped.emit(self.file_path)
                self.setStyleSheet("""QLabel { border: 2px dashed #aaaaaa;
                                           background-color: #f9f9f9;
                                           color: #555555;
                                           font-size: 14px; 
                                           padding: 20px;}""")

            event.acceptProposedAction()
        else:
            event.ignore()

    def clearFile(self):
        self.setText(self.default_text)
        self.file_path = None

    def sendtoverif (self,client):
            
            if self.file_path is not None:
                self.data.file_processed = self.character.name_processed
                self.data.file_processedpath = self.character.file_processed
                now = datetime.now()
                time = now.strftime('%Y%m%d-%H%M%S')
                subfolder_name = f"{client}-{time}"
                self.data.name = client
                
                
                subfolder_path = os.path.join(self.data.file_processedpath,subfolder_name)
                target = os.path.join(subfolder_path, os.path.basename(self.file_path))
                
                
                self.data.subfolder_name = subfolder_name
                self.data.file_name = self.file_path
                self.data.target = target
                self.data.subfolder_path = subfolder_path
                self.data.time = time
                self.data.file_loggerpath = self.character.file_loggerpath

                self.data.user_id = self.character.user_id
                self.data.folder_id = self.character.folder_id
                self.data.absolut_path = self.character.absolut_path

class ClearButton(QPushButton):
    def __init__(self,text = 'CLEAR', parent = None):
        super().__init__(text,parent)
        self.setStyleSheet("""
                           QPushButton { background-color: rgba(208, 19, 26,200);
                                    color: white;
                                    border: 2px solid darkred;
                                    border-radius: 10px; 
                                    padding: 8px 16px;
                           }
                           QPushButton:  hover {
                           background-color: rgba(208, 19, 26,200)
                           }
                           QPushButton: pressed {
                           background-color: rgba(208, 19, 26,200);}
                           """)
                        
        self.setMinimumSize(200,90)

class AcceptButton(QPushButton):
    def __init__(self,text = 'ACCEPT', parent = None):
        super().__init__(text,parent)
        self.setStyleSheet("""
                           QPushButton { background-color: rgba(0, 153, 65,200);
                                    color: white;
                                    border: 2px solid darkgreen;
                                    border-radius: 10px; 
                                    padding: 8px 16px;
                           }
                           QPushButton:  hover {
                           background-color: rgba(0, 153, 65,200)
                           }
                           QPushButton: pressed {
                           background-color: rgba(0, 153, 65,200);}
                           """)
        self.setMinimumSize(200,90)

#--------------------------------------------#

class AcceptButtonVerif(QPushButton):
    def __init__(self,text = 'ACCEPT', parent = None):
        super().__init__(text,parent)
        self.setStyleSheet("""
                           QPushButton { background-color: rgba(0, 153, 65,200);
                                    color: white;
                                    border: 2px solid darkgreen;
                                    border-radius: 10px; 
                                    padding: 8px 16px;
                           }
                           QPushButton:  hover {
                           background-color: rgba(0, 153, 65,200)
                           }
                           QPushButton: pressed {
                           background-color: rgba(0, 153, 65,200);}
                           """)

class CancelButtonVerif(QPushButton):
    def __init__(self,text = 'CANCEL', parent = None):
        super().__init__(text,parent)
        self.setStyleSheet("""
                           QPushButton { background-color: rgba(208, 19, 26,150);
                                    color: white;
                                    border: 2px solid darkred;
                                    border-radius: 10px; 
                                    padding: 8px 16px;
                           }
                           QPushButton:  hover {
                           background-color: rgba(208, 19, 26,150)
                           }
                           QPushButton: pressed {
                           background-color: rgba(208, 19, 26,150);}
                           """)

# =================================================================================================
# =                                SECTION: INFORMATIONS TRANSMISES
# =================================================================================================

class data : 
    "Return informations for verification"
    def __init__(self):
        self.subfolder_name = ''
        self.subfolder_path = ''
        self.name = '' 
        self.file_name = ''
        self.target = '' 
        self.file_processed = '' 
        self.file_processedpath = ''
        self.file_loggerpath = ''
        self.absolut_path = ''
        self.time = '' 
        self.folder_id = '' 
        self.user_id = '' 

  