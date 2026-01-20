# =================================================================================================
# =                                       SECTION: IMPORT
# =================================================================================================
import sys, os
import pymssql
import json
import pandas as pd
from rapidfuzz import fuzz, process 
from BT_main import UseInfo
from pathlib import Path
import numpy as np
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QWidget, QComboBox, QHBoxLayout,
                            QVBoxLayout, QPushButton, QLabel, QLineEdit, 
                            QListWidget, QMessageBox, QTableWidget, QTableWidgetItem,
                            QTabWidget, QFrame, QTableWidgetItem, QMenu,
                            QInputDialog,QGridLayout, QCheckBox
                            )

from PyQt6.QtCore import Qt , pyqtSignal, QPoint
from PyQt6.QtGui import QColor, QAction, QFont, QPalette
from Style import get_stylesheet
from Discussion_bloom import discussion
import logging, string
import os, shutil
from datetime import datetime


import re 

CONFIG = 'config.json'

# =================================================================================================
# =                                       SECTION: FOLDERS
# =================================================================================================


# =================================================================================================
# =                          SECTION: FENETRE DE VERIFICATION DES COLONNES
# =================================================================================================


class ColumnWidget(QWidget):
    """Forme une ComboBox lié avec un QLineEdit/QComboBox l'un au dessus de l'autre
            - le Combobox représente la colonne attendue
            - Le QLineEdit/QComboBox en dessous dépend de la Combobox et réprénsente un valeur à vouloir remplir automatiquement
            - QLineEdit ou QComboBox dépend de la valeur dans la Combobox"""
    
    def __init__(self, col_name,nkey, Lkey, Ckey,get_value_for_key, client_accounts , get_default_key_for_column,val_auto, parent = None):
        super().__init__(parent)
        self.col_name = col_name
        self.val_auto = val_auto 
        self.client_accounts = client_accounts
        self.nkey = nkey
        self.Lkey = Lkey
        self.Ckey = Ckey
        self.get_value_for_key = get_value_for_key
        self.get_default_key_for_column = get_default_key_for_column
        self.setStyleSheet(get_stylesheet())

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0,0,0,0)
        self.layout.setSpacing(2)

        self.combo_selector = QComboBox()
        self.combo_selector.addItems(self.nkey)
        

        default_key = self.get_default_key_for_column(self.col_name)

        if default_key in self.nkey:
            self.combo_selector.setCurrentText(default_key)
        else:
            self.combo_selector.setCurrentIndex(-1)
        
        
        self.dynamic_widget = None 

        self.create_dynamic_widget(self.combo_selector.currentText())
        
        self.layout.addWidget(self.dynamic_widget)

        
        self.combo_selector.currentTextChanged.connect(self.on_key_changed)

        self.layout.addWidget(self.combo_selector)
        if self.dynamic_widget:
            self.layout.addWidget(self.dynamic_widget)
       

        
    def get_mapping(self):
        """
        Trie les informations renseignées dans la fenêtre d'assignation des colonnes avant vérification

        """
        key = self.combo_selector.currentText()
        if not key or key == '':
            return None, None, None  
        
        value_widget = None 
        if isinstance(self.dynamic_widget, QLineEdit):
            value_widget = self.dynamic_widget.text().strip()
        elif isinstance(self.dynamic_widget, QComboBox):
            value_widget = self.dynamic_widget.currentText().strip()
        
        if value_widget:
            return key, "Manuel", value_widget
        else:
            return key, self.col_name, None

    def create_dynamic_widget(self,key):
        """ Crée le widget dynamique selon la selection"""
        
        if self.dynamic_widget:
            self.dynamic_widget.deleteLater()
            self.dynamic_widget = None
        
        if key in self.Lkey:
            self.dynamic_widget = QLineEdit()
            self.dynamic_widget.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
            self.dynamic_widget.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
            self.dynamic_widget.setFocus()
            if self.val_auto:
                self.dynamic_widget.setText(self.val_auto)
        
            
        elif key in self.Ckey:
            self.dynamic_widget = QComboBox()
            
            self.dynamic_widget.raise_()
            self.dynamic_widget.addItems(self.get_value_for_key(key,self.client_accounts))
            if self.val_auto:
                self.dynamic_widget.setCurrentText(self.val_auto)

        if self.dynamic_widget:
           
            self.dynamic_widget.setEnabled(True)
            self.layout.addWidget(self.dynamic_widget)
            
    def on_key_changed(self, new_key):

        self.create_dynamic_widget(new_key)  
    
    def setFixedWidth(self, width):
        super().setFixedWidth(width)
        if self.combo_selector:
            self.combo_selector.setFixedWidth(width)
        if self.dynamic_widget:
            self.dynamic_widget.setFixedWidth(width)

class ScrollSyncOverlay(QWidget):
    """Widget invisible qui place les ColumnWidgets sous les colonnes (complex pour pas grand)"""
    def __init__(self,table,column_widgets):
        super().__init__(table.parent())
        self.table = table
        self.column_widget = column_widgets
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setStyleSheet(get_stylesheet())
        self.raise_()
     
        header = table.horizontalHeader()
        header.sectionResized.connect(self.update_positions)
        table.viewport().installEventFilter(self)

        self.update_positions()

    def eventFilter(self, obj, event):
        """Quand la table bouge/redessine, or repositionne"""
        if obj is self.table.viewport():
            self.update_positions()
        return super().eventFilter(obj,event)
    
    def update_positions(self):
        """Positionne chaque ColumnWidget pile sous sa colonne."""
        header = self.table.horizontalHeader()
        scroll = self.table.horizontalScrollBar().value()
        y = self.table.geometry().bottom() + 5 
        
        for i, widget in enumerate(self.column_widget):
            if i >= self.table.columnCount():
                continue
            x = header.sectionPosition(i) - scroll + self.table.geometry().x()
            width = header.sectionSize(i)
            widget.setParent(self.parent())
            widget.setGeometry(x,y,width, widget.sizeHint().height())
            widget.raise_()
            widget.show()

class RulesIndicator(QWidget):
    """Widget des règles à utiliser pour le client"""
    def __init__(self,rules_dic,parent = None ):
        super().__init__(parent)
        self.rules_dic = rules_dic
        self.setStyleSheet(get_stylesheet())

        layout = QGridLayout()
        layout.setContentsMargins(15,15,15,15)
        layout.setHorizontalSpacing(10)
        layout.setSpacing(6)

        key_font = QFont("Aria",10)
        value_font = QFont("Aria", 10 )
        value_font.setBold(True)

        row = 0
        self.inputs = {}

        for key, value in self.rules_dic.items():
            key_label = QLabel(str(key))
            key_label.setFont(key_font)
            key_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            key_label.setStyleSheet("color: #555;")

            check_box = QCheckBox()
            check_box.setChecked(True)
            check_box.setStyleSheet("QCheckBox:: indicator {width: 14px; height: 14px;}")



            value_edit = QLineEdit()
            if value:
                value_edit.setText(value)
            value_edit.setFont(value_font)

            layout.addWidget(key_label,row,0)
            layout.addWidget(check_box, row , 1, alignment=Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(value_edit, row, 2 )

            self.inputs[key] = {"checkbox": check_box, "edit": value_edit}
            row +=1
            
        
        self.setLayout(layout)
        self.setWindowTitle("Indicateurs")
        self.setFixedSize(300, min(400, 40 + 25 * len(self.rules_dic)))

        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#f9f9f9"))
        self.setPalette(palette)
        self.setAutoFillBackground(True)
        
class DataFrameWidget(QWidget):
    validated = pyqtSignal(dict,dict,dict)
    def __init__(self,fclose,client_ulId,df,nkey,Lkey,Ckey,get_value_for_key, get_default_key_for_column,client_accounts, memory, memory_comparison_client,dic_rules,mapping_currency,mapping_side,mapping_type,mapping_validitytype,parent = None):
        super().__init__(parent)
        
        self.fclose = fclose 
        self.force = False 
        self.df = df
        self.nkey = [""] + nkey + ["Date"]
        self.Lkey = Lkey
        self.Ckey = Ckey
        self.get_value_for_key = get_value_for_key
        self.get_default_key_for_column = get_default_key_for_column
        self.client_accounts = client_accounts
        self.extra_column_widget = []
        self.column_widget = []
        self.memory = memory
        self.list_data = ["ValidityType_data","OrderType_data"]
        self.list_data_corres = ["ValidityType", "OrderType"]
        self.memory_comparison_client = memory_comparison_client
        self.client_ulId = client_ulId
        self.rules_dic = dic_rules
        
        self.mapping_currency = mapping_currency
        self.mapping_side = mapping_side
        self.mapping_type = mapping_type
        self.mapping_validitytype = mapping_validitytype
        

        self.resize(1900,800)
      
       
        self.setStyleSheet(get_stylesheet())
        

        main_layout = QVBoxLayout(self)


        self.table = QTableWidget()

        if self.memory_comparison_client:
            # affiche le tableau en jaune lorsque le format est reconnu 
            self.table.setStyleSheet("background-color: #ffffe0")

        self.table.setMaximumHeight(150)
        n_rows = min(3,len(self.df))
        n_cols =len(self.df.columns)
        self.table.setRowCount(n_rows)
        self.table.setColumnCount(n_cols)
        self.table.setHorizontalHeaderLabels([str(c) for c in self.df.columns])

        for i in range(n_rows):
            for j,col in enumerate(self.df.columns):
                text = str(self.df.iloc[i,j]).strip()
                if text =='nan' or text == '?':
                    text = ''
                item = QTableWidgetItem(text)
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                self.table.setItem(i,j,item)

        main_layout.addWidget(self.table)

        
        for col_name in df.columns:
            if self.memory_comparison_client :
                reverse_memory = {}
                for k,v in self.memory.items():
                    if k == 'Comment' and v is not None:
                        if '#' in v:
                            # Lorsqu'un commentaire est sur plusieurs colonnes, il est noté 'col1#col2' en mémoire, il faut donc séparé les deux colonnes pour extraire les données. 
                            list_comment = v.split('#')
                        else :
                            list_comment = [v]
                        for i in list_comment:
                            reverse_memory.setdefault(i,k)
                    else:
                        reverse_memory.setdefault(v,k)
                  
                    default_key = reverse_memory.get(col_name)
                                         
            else : 
                default_key = self.get_default_key_for_column(col_name)
                if default_key == None:
                    default_key = self.no_titled(col_name)
                
            
            self.column_widget += [ColumnWidget(col_name,self.nkey,Lkey,Ckey,get_value_for_key, self.client_accounts, get_default_key_for_column = lambda _: default_key, val_auto= None)]

        self.overlay = ScrollSyncOverlay(self.table, self.column_widget)
            
        self.overlay.update_positions()
        
        
        # Separateur visuel 
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(sep)

        # Ligne extra 

        self.extra_layout = QHBoxLayout()
        main_layout.addLayout(self.extra_layout)


        self.layout_extra = QHBoxLayout()
        # Bouton ajouter extra

        self.button_add_column = QPushButton("Ajouter une colonne")
        self.button_add_column.clicked.connect(lambda : self.add_extra_column_widget(None,None))
        self.button_add_column.setFixedHeight(50)
        self.button_add_column.setFixedWidth(150)
        self.layout_extra.addWidget(self.button_add_column)

        # Ajout automatique 
        if self.memory_comparison_client:
           
            for idx, data in enumerate(self.list_data):
                col = list(self.memory.keys())[15 + idx]
                val = self.memory[col]
                col_corres = self.list_data_corres[idx]
                if val: 
                    self.add_extra_column_widget(col_corres,val)
                

        
        # Bouton supprimer extra
        self.button_delete_extra = QPushButton("Supprimer une colonne")
        self.button_delete_extra.clicked.connect(self.remove_last_extra_column_widget)
        self.button_delete_extra.setFixedHeight(50)
        self.button_delete_extra.setFixedWidth(150)
        self.layout_extra.addWidget(self.button_delete_extra)
        main_layout.addLayout(self.layout_extra)

        
        layout_bot = QHBoxLayout()
        # ADD MAPPING 
        self.correctionbutton = CorrectionButton(self.client_ulId)
        self.correctionbutton.clicked.connect(self.correctionbutton.open_correction)
        layout_bot.addWidget(self.correctionbutton, alignment=Qt.AlignmentFlag.AlignLeft)

        # Rules 

        self.rules_indicator = RulesIndicator(self.rules_dic)
        layout_bot.addWidget(self.rules_indicator)
        

        # Valider 
        self.valider_btn = QPushButton("Valider")
        self.valider_btn.setFixedHeight(50)
        self.valider_btn.setFixedWidth(150)
        layout_bot.addWidget(self.valider_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.valider_btn.clicked.connect(self.valider_action)
        main_layout.addLayout(layout_bot)
    
    def no_titled(self,val):
        """ Entrée : 
                    - val : Nom d'une colonne
            Objectif : Determiner le type de colonne lorsqu'elle ne possède pas de titre"""
        
        if 'Column' in val or 'Unnamed' in val :
            data = self.df[val].iloc[0] 

            

            liste_mapping_side = []
            
            for i in self.mapping_side.values():
                liste_mapping_side.extend(i)
            
            liste_mapping_type = []
            for i in self.mapping_type.values():
                liste_mapping_type.extend(i)

            liste_mapping_validitytype = []
            for i in self.mapping_validitytype.values():
                liste_mapping_validitytype.extend(i)
            
            if data in liste_mapping_side:
                return 'Side'
            
            if data in liste_mapping_type:
                return 'Ordertype'
            
            if data in liste_mapping_validitytype:
                return 'ValidityType'
        
            if isinstance(data,np.integer):
                if data >= 5000000000 and data < 6000000000:
                    return 'Account'
                else:
                    return 'Quantity'
                
            if isinstance(data,str):
                if len(data.strip()) == 12:
                    return 'isin'
                if len(data.strip()) == 3:
                    return 'Currency'
            
                if len(data.strip()) == 4:
                    return 'Exchange'
            
        else:

            return None 


    def add_extra_column_widget(self, col, val):
        """Ajoute dynamiquement un ColumnWidget"""
        
        if val : 
            col_widget = ColumnWidget(
                                    col_name = "(Nouvelle colonne)",
                                    nkey = self.nkey,
                                    Lkey = self.Lkey,
                                    Ckey = self.Ckey,
                                    get_value_for_key=self.get_value_for_key,
                                    get_default_key_for_column= lambda _: col,
                                    client_accounts = self.client_accounts,
                                    val_auto= val)
        else : 
            col_widget = ColumnWidget(
                                    col_name = "(Nouvelle colonne)",
                                    nkey = self.nkey,
                                    Lkey = self.Lkey,
                                    Ckey = self.Ckey,
                                    get_value_for_key=self.get_value_for_key,
                                    get_default_key_for_column= lambda _:" ",
                                    client_accounts = self.client_accounts,
                                    val_auto= None)
        self.extra_layout.addWidget(col_widget)
        self.extra_column_widget.append(col_widget)

    def remove_last_extra_column_widget(self):
        """ Supprime le dernier ColumnWidget ajouté"""
        if not self.extra_column_widget:
            return

        last_widget = self.extra_column_widget.pop()
        self.extra_layout.removeWidget(last_widget)

        last_widget.setParent(None)
        last_widget.deleteLater()

    

    def valider_action(self):
        """Vérifie les colonnes manquantes et les doublons """
        
        result = {}
        dic_man = {}
        double_select = []
        no_select = []
        ex_cur_bloom_select = 0
        list_comment = []

        for col_widget in self.column_widget:
            key, value_1, value_2 = col_widget.get_mapping()
            is_valid_type = (((key == 'ValidityType' or key == 'OrderType') and 'ValidityType & OrderType' in result) or (key == 'ValidityType & OrderType' and ('ValidityType' in result or 'OrderType' in result)))
            if key and value_2:
                if key not in result and key != 'Comment' and not is_valid_type:
                    result[key] = ["Manuel"]
                    dic_man[key] = value_2
                elif key in result or is_valid_type : 
                    double_select.append(key)
               
                    
            elif key and not value_2:
                if key not in result and key != 'Comment' and key != 'ValidityType & OrderType' and not is_valid_type:
                    result[key] = [value_1]
                elif key in result and is_valid_type : 
                    double_select.append(key)
                elif key == 'Comment':
                    list_comment.append(value_1)
                elif key == 'ValidityType & OrderType':
                    result['ValidityType'] = [value_1]
                    result['OrderType'] = [value_1]
                    result['ValidityType & OrderType'] = [value_1]
                

                
        for col_extra_widget in self.extra_column_widget:
            key, value_1, value_2 = col_extra_widget.get_mapping()
            is_valid_type = (((key == 'ValidityType' or key == 'OrderType') and 'ValidityType & OrderType' in result) or (key == 'ValidityType & OrderType' and ('ValidityType' in result or 'OrderType' in result)))
            
            if key and value_2:
                if key not in result and key != 'Comment' and not is_valid_type:
                    result[key] = ["Manuel"]
                    dic_man[key] = value_2
                elif key in result and is_valid_type: 
                    double_select.append(key)

        if list_comment:
            result['Comment'] = list_comment
                    

        
        if 'Bloomberg Code' in result:
            ex_cur_bloom_select = 0
        
        for colonne in ['isin','Side','Quantity','OrderType','ValidityType']:
            if colonne not in result:
                if (colonne == 'ValidityType' or colonne == 'OrderType') and 'ValidityType & OrderType' in result:
                    continue
                elif colonne == 'ValidityType' and 'Date' in result :
                    continue
                elif colonne == 'isin' and  'Bloomberg Code' in result:
                    continue 
                else:
                    no_select.append(colonne)
        
        if double_select != [] or no_select != [] or ex_cur_bloom_select == 1 :
            
            msg = QMessageBox()
            message =''

            if double_select:
                message += f"\n Colonnes sélectionnées deux fois : \n "

                for col in double_select:
                    message += f" Colonne : {col} \n"
            
            if no_select:
                message += f"\n Colonnes non sélectionnées : \n"

                for col in no_select:
                    message += f"Colonne : {col}\n"
            
           

            # Créer une fenêtre d'avertissement non-modale
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("Colonne Manquante")
            msg.setText(message)
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.setModal(False)  # Non-modal = ne bloque pas
            
            msg.exec()
        
        else: 
            inputs = {}
            for  key in self.rules_indicator.inputs.keys():
                inter = {}
                if self.rules_indicator.inputs[key]['checkbox'].isChecked():     
                    inter['checkbox'] = True
                    text = self.rules_indicator.inputs[key]['edit'].text().strip()
                    if text == '':
                        inter['edit'] = None
                    else: 
                        inter['edit'] = text

                else:
                    inter['checkbox'] = False
                    inter['edit'] = None
                inputs[key] = inter

            self.validate_selection(result,dic_man,inputs)

    
    def validate_selection(self,result,dic_man,inputs):
        self.validated.emit(result,dic_man,inputs)
        self.force = True 
        self.close()



    def closeEvent(self,event):
        if not self.force :
            QApplication.closeAllWindows()
            QApplication.quit()
            event.accept()
            self.fclose.close()



# =================================================================================================
# =                                 SECTION: IDENTIFICATION DES COLONNES 
# =================================================================================================
        
def get_default_key_for_column (col_name):
    """Cherche si une colonne du fichier est présente dans le mapping des colonnes et retourne la colonne attendu si trouvée """
    conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
    cursor = conn.cursor()
    cursor.execute("SELECT colonne_attendue, colonne_variante FROM TBL_BTRADE_COLNAME")

    mapping = {}
    for attendu, variante in cursor.fetchall():
        mapping[variante] = attendu

    conn.close()
    cursor.close()
    return mapping.get(col_name.strip(),None)
    
def get_value_for_key(key,client_accounts):
    """ Offre un remplissage limité pour certaines avec colonnes à l'aide d'un combobox"""
    
    mapping = {
        "Side" : [" ","Buy", "Sell"],
        "OrderType" : [" ","Limit", "Market", "Stop","StopLimit"],
        "ValidityType": [" ","Day","GTD","GTC","AtClose","AtOpen"],
    }
    mapping["Account"] = [" "] + client_accounts
    
    return mapping.get(key,[])



# =================================================================================================
# =                                      SECTION: CLASS AUDIT 
# =================================================================================================


class AuditLogger:
    """Piste d'audit """
    def __init__(self,log_dir ,level = logging.INFO):

        os.makedirs(log_dir, exist_ok = True)
        date = datetime.now().strftime('%Y-%m-%d')
        log_path = 'Log_'+date
        log_dir = os.path.join(log_dir,log_path)
        print('log_path',log_path)
        os.makedirs(log_dir,exist_ok=True)

        self.log_file = os.path.join(log_dir,f"audit_{datetime.now().strftime('%Y%m%d-%H%M%S')}.log")

        # Configuration du logger 
        self.logger = logging.getLogger("AuditLogger")
        self.logger.setLevel(level)

        # Eviter de dupliquer les handlers si plusieurs instances
        if not self.logger.handlers:
            file_handler = logging.FileHandler(self.log_file,mode = 'a', encoding = 'utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

    def info(self,message, *args):
        self.logger.info(message,*args)
    
    def warning(self, message, *args):
        self.logger.warning(message,*args)

    def error(self,message,*args):
        self.logger.error(message,*args)

class forceclose:
    """Force l'arret du programme en cas de fermeture d'une fenetre: 
            - Supprime le log 
            - Supprime le fichier crée """
    
    def __init__(self,log,subfolder_path):
        self.log = log
        self.log_name = log.log_file
        self.dossier = subfolder_path

    def close(self):
        for handler in self.log.logger.handlers[:]:
            if isinstance(handler,logging.FileHandler):
                handler.close()
                self.log.logger.removeHandler(handler)
        os.remove(self.log_name)
        
        shutil.rmtree(self.dossier)
        sys.exit(0)


# =================================================================================================
# =                                    SECTION: CLIENT ULID SECURITE
# =================================================================================================


class SelectionWindow(QWidget):
    """Permet de séléctionner un UlId parmits ceux de la liste et l'envoie pour l'allouer à self.ClientUlId """
    selection_ready = pyqtSignal(str)  # Signal pour retourner l'élément sélectionné
    
    def __init__(self, items_list, fclose,title="Sélection", parent=None):
        super().__init__(parent)
        self.fclose = fclose
        self.force = False
        self.items_list = items_list
        self.selected_item = None
        self.window_title = title
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle(self.window_title)
        self.setMinimumSize(300, 150)
        
        main_layout = QVBoxLayout()
        
        # Titre
        title_label = QLabel("Sélectionnez un Ul_Id :")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 10px;")
        main_layout.addWidget(title_label)
        
        # ComboBox avec la liste
        self.combo_box = QComboBox()
        self.combo_box.addItems(self.items_list)
        if self.items_list:  # Sélectionner le premier élément par défaut
            self.combo_box.setCurrentIndex(0)
        main_layout.addWidget(self.combo_box)
        
        # Bouton Valider
        validate_btn = QPushButton("Valider")
        validate_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; margin-top: 10px; }")
        validate_btn.clicked.connect(self.validate_selection)
        main_layout.addWidget(validate_btn)
        
        self.setLayout(main_layout)
    
    def validate_selection(self):
        """Valider la sélection et émettre le signal"""
        self.selected_item = self.combo_box.currentText()
        self.selection_ready.emit(self.selected_item)
        self.force = True 
        self.close()
    
    def get_selected_item(self):
        """Retourner l'élément sélectionné"""
        return self.selected_item

    def closeEvent(self,event):
        if not self.force:
            QApplication.closeAllWindows()
            QApplication.quit()
            event.accept()

            self.fclose.close()

class ClientIdInputWindow(QWidget):
    """Demande de saisir directement l'UlId du client"""

    client_id_ready = pyqtSignal(str)  # Signal pour retourner le Client Ul_Id saisi
    
    def __init__(self, fclose, title="Saisie Client Ul_Id", parent=None):
        """
        Fenêtre de saisie pour le Client Ul_Id
        
        Args:
            title (str): Titre de la fenêtre
            parent: Widget parent (optionnel)
        """
        super().__init__(parent)
        self.client_id = None
        self.window_title = title
        self.fclose = fclose
        self.force = False
        self.init_ui()
    
    def init_ui(self):
        """Initialiser l'interface utilisateur"""
        self.setWindowTitle(self.window_title)
        self.setMinimumSize(350, 150)
        
        main_layout = QVBoxLayout()
        
        # Titre
        title_label = QLabel("Saisissez le Client Ul_Id :")
        title_label.setStyleSheet("font-weight: bold; font-size: 14px; margin-bottom: 10px;")
        main_layout.addWidget(title_label)
        
        # EditLine pour saisir le Client Ul_Id
        self.line_edit = QLineEdit()
        self.line_edit.setPlaceholderText("Entrez le Client Ul_Id...")
        self.line_edit.setStyleSheet("padding: 8px; font-size: 12px; border: 2px solid #ccc; border-radius: 4px;")
        # Permet de valider en appuyant sur Entrée
        self.line_edit.returnPressed.connect(self.validate_input)
        main_layout.addWidget(self.line_edit)
        
        # Bouton Valider
        validate_btn = QPushButton("Valider")
        validate_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; margin-top: 10px; }")
        validate_btn.clicked.connect(self.validate_input)
        main_layout.addWidget(validate_btn)
        
        self.setLayout(main_layout)
        
        # Donner le focus à l'EditLine
        self.line_edit.setFocus()
    
    def validate_input(self):
        """Valider la saisie et émettre le signal"""
        self.client_id = self.line_edit.text().strip()
        
        if not self.client_id:
            QMessageBox.warning(self, "Erreur", "Veuillez saisir un Client Ul_Id valide.")
            return
        
        self.client_id_ready.emit(self.client_id)
        self.force = True 
        self.close()
    
    def get_client_id(self):
        """Retourner le Client Ul_Id saisi"""
        return self.client_id

    def closeEvent(self,event):
        if not self.force:
            QApplication.closeAllWindows()
            QApplication.quit()
            event.accept()
            self.fclose.close()


def get_client_accounts(client_UlID):
    """ Forme une la liste des Ul_Id des comptes du client à partir de la table TBLACCOUNTS et une liste de dictionnaire:
        - Chaque disctionnaire correspond à une ligne de compte du client
        - Chaque dictionnaire prend en clés les colonnes  Ul_Id, AccountId, Description, ExternalIds de la table TBLACCOUNTS """

    client_accounts = []
    conn = pymssql.connect(
            server = '',
            user = '',
            password= '',
            database=''
        )
    cursor = conn.cursor()
    cursor.execute("SELECT Ul_Id FROM TBLACCOUNTS WHERE ClientId = %s AND Ul_Id NOT LIKE '%_copy' AND Active != 'false' ", (client_UlID,))
    row = cursor.fetchall()
    for item in row:
        client_accounts.append(str(item[0]))
    l = []
   
    cursor.execute ("SELECT Ul_Id, AccountId, Description, ExternalIds FROM TBLACCOUNTS WHERE ClientId = %s AND Ul_Id NOT LIKE '%_copy' AND Active != 'false'", (client_UlID,))
    row2 = cursor.fetchall()
    
    for accounts in row2 : 
        d = {}
        for i in accounts:
            d[i] = str(accounts[0]).strip().upper()
        l.append(d)
        

    return client_accounts, l
    

# =================================================================================================
# =                                  SECTION: AJOUT BASE DE DONNEE SQL 
# =================================================================================================


class MappingApp(QWidget):
    """Fenetre permettant d'ajouter des informations dans la base de donnée SQL """
    def __init__(self,audit,colonne_name,client_name):
        super().__init__()
        self.audit = audit 
        self.conn =  pymssql.connect(
        server = '',
        user = "",
        password = '',
        database = '')

        self.setWindowTitle("Fenetre Principal")
        self.setFixedSize(500,300)
        self.client_name = client_name
        layout = QVBoxLayout()

        self.tabs = QTabWidget()

        # FIRST WINDOW 

        self.tab1 = QWidget()
        self.tab1.setWindowTitle("Ajouter une variante de colonne")
        self.tab1_layout  = QVBoxLayout()
        self.tab1.label_attendue = QLabel("Colonne attendue:")
        self.tab1_layout.addWidget(self.tab1.label_attendue)

        self.tab1.combo_attendue = QComboBox()
        self.tab1.combo_attendue.setEditable(True)
        self.tab1.combo_attendue.addItems(colonne_name)
        self.tab1.combo_attendue.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.tab1_layout.addWidget(self.tab1.combo_attendue)

        self.tab1.label_variante = QLabel("Variante : ")
        self.tab1_layout.addWidget(self.tab1.label_variante)
        self.tab1.input_variante = QLineEdit()
        self.tab1_layout.addWidget(self.tab1.input_variante)

        self.tab1.btn_ajouter = QPushButton("Ajouter")
        self.tab1_layout.addWidget(self.tab1.btn_ajouter)
        self.tab1.btn_ajouter.clicked.connect(self.ajouter_click_column)

        self.tab1.setLayout(self.tab1_layout)


        # SECOND WINDOW 


        self.tab2 = QWidget()
        self.tab2.setWindowTitle("Ajouter une variante de currency")
        self.tab2.searchbox = SearchBox(self.conn)

        self.tab2.searchbox.results_ready.connect(self.update_list)

        self.tab2.list_widget = QListWidget()
        self.tab2.list_widget_value = None
        self.tab2.list_widget.itemClicked.connect(self.on_item_clicked)

        self.tab2.label_variante = QLabel("Variante : ")
        self.tab2.input_variante = QLineEdit()
        

        

        self.tab2.btn_ajouter = QPushButton("Ajouter")
        self.tab2_layout  = QVBoxLayout()
        self.tab2_layout.addWidget(self.tab2.searchbox)
        self.tab2_layout.addWidget(self.tab2.list_widget)
        self.tab2_layout.addWidget(self.tab2.label_variante)
        self.tab2_layout.addWidget(self.tab2.input_variante)
        self.tab2_layout.addWidget(self.tab2.btn_ajouter)

        self.tab2.btn_ajouter.clicked.connect(self.ajouter_click_currency)
        self.tab2.list_widget.itemClicked.connect(self.item_selected)

        self.tab2.setLayout(self.tab2_layout)
        

        # THIRD WINDOW 

        self.tab3 = QWidget()
        self.tab3.setWindowTitle("Ajouter une variante de Side")
        self.tab3_layout  = QVBoxLayout()
        self.tab3.label_attendue = QLabel("Side attendue:")
        self.tab3_layout.addWidget(self.tab3.label_attendue)

        self.tab3.combo_attendue = QComboBox()
        self.tab3.combo_attendue.setEditable(True)
        self.tab3.combo_attendue.addItems(self.get_distinct_side())

        self.tab3.combo_attendue.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.tab3_layout.addWidget(self.tab3.combo_attendue)

        self.tab3.label_variante = QLabel("Variante : ")
        self.tab3_layout.addWidget(self.tab3.label_variante)
        self.tab3.input_variante = QLineEdit()
        self.tab3_layout.addWidget(self.tab3.input_variante)

        self.tab3.btn_ajouter = QPushButton("Ajouter")
        self.tab3_layout.addWidget(self.tab3.btn_ajouter)
        self.tab3.btn_ajouter.clicked.connect(self.ajouter_click_side)

        self.tab3.setLayout(self.tab3_layout)

        # FOURTH WINDOW 

        self.tab4 = QWidget()
        self.tab4.setWindowTitle("Ajouter une variante de Side")
        self.tab4_layout  = QVBoxLayout()
        self.tab4.label_attendue = QLabel("Side attendue:")
        self.tab4_layout.addWidget(self.tab4.label_attendue)

        self.tab4.combo_attendue = QComboBox()
        self.tab4.combo_attendue.setEditable(True)
        self.tab4.combo_attendue.addItems(self.get_distinct_type())

        self.tab4.combo_attendue.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.tab4_layout.addWidget(self.tab4.combo_attendue)

        self.tab4.label_variante = QLabel("Variante : ")
        self.tab4_layout.addWidget(self.tab4.label_variante)
        self.tab4.input_variante = QLineEdit()
        self.tab4_layout.addWidget(self.tab4.input_variante)

        self.tab4.btn_ajouter = QPushButton("Ajouter")
        self.tab4_layout.addWidget(self.tab4.btn_ajouter)
        self.tab4.btn_ajouter.clicked.connect(self.ajouter_click_type)

        self.tab4.setLayout(self.tab4_layout)

        # FIFTH WINDOW 

        self.tab5 = QWidget()
        self.tab5.setWindowTitle("Ajouter une variante de ValidityType")
        self.tab5_layout  = QVBoxLayout()
        self.tab5.label_attendue = QLabel("ValidityType attendue:")
        self.tab5_layout.addWidget(self.tab5.label_attendue)

        self.tab5.combo_attendue = QComboBox()
        self.tab5.combo_attendue.setEditable(True)
        self.tab5.combo_attendue.addItems(self.get_distinct_validitytype())

        self.tab5.combo_attendue.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.tab5_layout.addWidget(self.tab5.combo_attendue)

        self.tab5.label_variante = QLabel("Variante : ")
        self.tab5_layout.addWidget(self.tab5.label_variante)
        self.tab5.input_variante = QLineEdit()
        self.tab5_layout.addWidget(self.tab5.input_variante)

        self.tab5.btn_ajouter = QPushButton("Ajouter")
        self.tab5_layout.addWidget(self.tab5.btn_ajouter)
        self.tab5.btn_ajouter.clicked.connect(self.ajouter_click_validitytype)

        self.tab5.setLayout(self.tab5_layout)


        # SIXTH WINDOW 

        self.tab6 = QWidget()
        self.tab6.setWindowTitle("Ajouter une variante de Règle")
        self.tab6_layout  = QVBoxLayout()
        self.tab6.label_attendue = QLabel("Règle :")
        self.tab6_layout.addWidget(self.tab6.label_attendue)

        self.tab6.combo_attendue = QComboBox()
        self.tab6.combo_attendue.setEditable(True)
        self.tab6.combo_attendue.addItems(self.get_distinct_rules())

        self.tab6.combo_attendue.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.tab6_layout.addWidget(self.tab6.combo_attendue)

        self.tab6.label_variante = QLabel("Code : ")
        self.tab6_layout.addWidget(self.tab6.label_variante)
        self.tab6.input_variante = QLineEdit()
        self.tab6_layout.addWidget(self.tab6.input_variante)

        self.tab6.btn_ajouter = QPushButton("Ajouter")
        self.tab6_layout.addWidget(self.tab6.btn_ajouter)
        self.tab6.btn_ajouter.clicked.connect(self.ajouter_click_rules)

        self.tab6.setLayout(self.tab6_layout)

        #

        self.tabs.addTab(self.tab1,"Add column")
        self.tabs.addTab(self.tab2, 'Add currency')
        self.tabs.addTab(self.tab3,'Add Side')
        self.tabs.addTab(self.tab4,'Add Type')
        self.tabs.addTab(self.tab5,'Add ValidityType')
        self.tabs.addTab(self.tab6,'Add Rule')
        
        layout.addWidget(self.tabs)
        self.setLayout(layout)
        

    # Fourni les possibilités de remplissage de la base sql
    def get_distinct_side(self):
        query = "SELECT DISTINCT side_attendue FROM TBL_BTRADE_SIDE"
        cursor = self.conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return [row[0]for row in rows if row[0] is not None ]
    
    def get_distinct_type(self):
        query = "SELECT DISTINCT type_attendue FROM TBL_BTRADE_TYPE"
        cursor = self.conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return [row[0]for row in rows if row[0] is not None ]

    def get_distinct_validitytype(self):
        query = "SELECT DISTINCT validitytype_attendue FROM TBL_BTRADE_VALIDITYTYPE"
        cursor = self.conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return [row[0]for row in rows if row[0] is not None ]

    def get_distinct_bloomberg(self):
        query = "SELECT DISTINCT bloomberg_attendue FROM TBL_BTRADE_BLOOMBERG"
        cursor = self.conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cursor.close()
        return [row[0]for row in rows if row[0] is not None ]
    
    def get_distinct_rules(self):
        query = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TBL_BTRADE_CLIENTMEMORY' ORDER BY ORDINAL_POSITION OFFSET 18 ROWS"
        cursor = self.conn.cursor()
        cursor.execute(query)
        rows = cursor.fetchall() 
        cursor.close()
        return [row[0]for row in rows if row[0] is not None ]
    

    def item_selected(self,item):
        self.tab2.attendue = item.text()
        
    def ajouter_click_column(self):
        col_attendue = self.tab1.combo_attendue.currentText().strip()
        variante = self.tab1.input_variante.text()

        

        if not col_attendue or not variante:
            QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
            return 
        
        sucess = add_variante_column (col_attendue,variante, self.conn)
        if sucess:
            QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {col_attendue} ")
            self.audit.info(f"Ajouté : {variante} à {col_attendue}")
            self.tab3.input_variante.clear()
            

        else:
            QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")

    def ajouter_click_currency(self):
            cur_attendue = self.tab2.attendue
            variante = self.tab2.input_variante.text()

            if not cur_attendue or not variante:
                QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
                return 
            
            sucess = add_variante_currency (cur_attendue,variante, self.conn)
            if sucess:
                QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {cur_attendue} ")
                self.audit.info(f"Ajouté : {variante} à {cur_attendue} ")
                self.tab2.input_variante.clear()
                

            else:
                QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")

    def ajouter_click_side(self):
        side_attendue = self.tab3.combo_attendue.currentText().strip()
        variante = self.tab3.input_variante.text()

        

        if not side_attendue or not variante:
            QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
            return 
        
        sucess = add_variante_side (side_attendue,variante, self.conn)
        if sucess:
            QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {side_attendue} ")
            self.audit.info(f"Ajouté : {variante} à {side_attendue} ")
            self.tab3.input_variante.clear()
            

        else:
            QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")

    def ajouter_click_type(self):
        type_attendue = self.tab4.combo_attendue.currentText().strip()
        variante = self.tab4.input_variante.text()

        

        if not type_attendue or not variante:
            QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
            return 
        
        sucess = add_variante_type (type_attendue,variante, self.conn)
        if sucess:
            QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {type_attendue} ")
            self.tab4.input_variante.clear()
            self.audit.info(f"Ajouté : {variante} à {type_attendue}")
            

        else:
            QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")

    def ajouter_click_validitytype(self):
        validitytype_attendue = self.tab5.combo_attendue.currentText().strip()
        variante = self.tab5.input_variante.text()

        

        if not validitytype_attendue or not variante:
            QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
            return 
        
        sucess = add_variante_validitytype (validitytype_attendue,variante, self.conn)
        if sucess:
            QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {validitytype_attendue} ")
            self.audit.info(f"Ajouté : {variante} à {validitytype_attendue} ")
            self.tab5.input_variante.clear()
            

        else:
            QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")
 
    def ajouter_click_rules(self):
        regle_attendue = self.tab6.combo_attendue.currentText().strip()
        variante = self.tab6.input_variante.text()

        

        if not regle_attendue or not variante:
            QMessageBox.warning(self, "Erreur", "Les deux champs sont obligatoires ! ")
            return 
        
        sucess = add_rule_code(regle_attendue,variante, self.conn, self.client_name)
        if sucess:
            QMessageBox.information(self,'Succes',f"Ajouté : {variante} à {regle_attendue} ")
            self.audit.info(f"Ajouté : {variante} à {regle_attendue} ")
            self.tab6.input_variante.clear()
            

        else:
            QMessageBox.critical(self,"Erreur","Impossible d'ajouter la variante.")

    def on_item_clicked(self,item):
        self.tab2.list_widget_value = item.text()

    def update_list(self,items):
        self.tab2.list_widget.clear()
        self.tab2.list_widget.addItems(items)

class SearchBox(QLineEdit):
    results_ready = pyqtSignal(list)

    def __init__(self, conn, parent = None):
        super().__init__(parent)

        self.textChanged.connect(self.on_text_changed)
        
        
        self.cursor = conn.cursor()
        self.setPlaceholderText("Research Currency :")


    def on_text_changed(self,text):
        if text.strip():
            results  = self.search_in_db(text)
            self.results_ready.emit(results)
        else:
            self.results_ready.emit([])
        
    def search_in_db(self,keyword):
        """Cherche le mot clé"""
        table_name = 'TBL_BTRADE_CURRENCY'
        column_name = 'currency_attendue'
        if not self.cursor:
            return []
        try:
            
            query = f"""
                SELECT [currency_attendue] 
                FROM [{table_name}] 
                WHERE {column_name} LIKE %s
                """

            param = ('%' + keyword + '%')
            self.cursor.execute(query,param)

            rows = self.cursor.fetchall()
            return [row[0] for row in rows]
        except Exception as e:
            import traceback
            print ("Erreur SQL :", e)
            traceback.print_exc()
            return []                



    # Remplissage de la base de donnée 

def add_variante_column(colonne_attendue,colonne_variante, conn):
    cursor = conn.cursor()
    try: 
        colonne_variante = colonne_variante.strip()
        cursor.execute("INSERT INTO TBL_BTRADE_COLNAME (colonne_attendue, colonne_variante) VALUES (%s,%s)",
                        (colonne_attendue,colonne_variante.strip())
        )
        cursor.execute("select * from TBL_BTRADE_COLNAME where colonne_attendue = 'isin'")
        conn.commit()

        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False

def add_variante_currency(currency_attendue,currency_variante, conn):
    
    cursor = conn.cursor()
    try: 
        currency_variante = currency_variante.strip()
        cursor.execute("INSERT INTO TBL_BTRADE_CURRENCY (currency_attendue, currency_variante) VALUES (%s,%s)",
                        (currency_attendue,currency_variante.strip())
        )
        conn.commit()

        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False

def add_variante_side(side_attendue,side_variante, conn):
    cursor = conn.cursor()
    try: 
        side_variante = side_variante.strip()
        cursor.execute("INSERT INTO TBL_BTRADE_SIDE (side_attendue, side_variante) VALUES (%s,%s)",
                        (side_attendue,side_variante.strip())
        )
        conn.commit()
        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False
    finally:
       pass

def add_variante_type(type_attendue,type_variante, conn):
    if type_variante == 'Type' or type_variante == 'type' or type_variante == 'TYPE':
        QMessageBox( title = 'Erreur', text = f" Erreur : {type_variante} ne peut pas être ajouté " ) 
        return 
    cursor = conn.cursor()
    try: 
        type_variante = type_variante.strip()
        cursor.execute("INSERT INTO TBL_BTRADE_TYPE (type_attendue, type_variante) VALUES (%s,%s)",
                        (type_attendue,type_variante.strip())
        )
        conn.commit()
        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False
    finally:
       pass

def add_variante_validitytype(validitytype_attendue,validitytype_variante, conn):
    if validitytype_variante == 'Type' or validitytype_variante == 'type' or validitytype_variante == 'TYPE':
        QMessageBox( title = 'Erreur', text = f" Erreur : {validitytype_variante} ne peut pas être ajouté " ) 
        return 
    cursor = conn.cursor()
    try: 
        validitytype_variante = validitytype_variante.strip()
        cursor.execute("INSERT INTO TBL_BTRADE_VALIDITYTYPE (validitytype_attendue, validitytype_variante) VALUES (%s,%s)",
                        (validitytype_attendue,validitytype_variante.strip())
        )
        conn.commit()
        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False
    finally:
       pass

def add_rule_code (rule , code, conn, client_name):
    """Ajout de règle en mémoire"""
    cursor = conn.cursor()
    try:
        code = code.strip()
        query = f"update TBL_BTRADE_CLIENTMEMORY set  %s = %s where ClientName = %s "
        params = [rule,code,client_name]
        cursor.execute(query,params)
        conn.commit()
        return True
    except Exception as e:
        print("Erreur SQL :",e)
        return False
    finally:
       pass


# =================================================================================================
# =                                   SECTION: MISE EN MEMOIRE DU CLIENT 
# =================================================================================================


class Demande_mise_en_memoire(QWidget):
    """Propose de mettre en mémoire les modification ou non, retourne un booleen"""
    sauvegarde = pyqtSignal(bool)
    def __init__(self,parent = None):
        super().__init__(parent)
        self.fclose = fclose
        self.force = False 
        self.init_ui()

    def init_ui(self):
        # Configuration de la fenêtre
        self.setWindowTitle("Memory")
        self.setMinimumSize(50, 100)  # Augmenté pour accommoder les nouvelles colonnes
        self.setStyleSheet(get_stylesheet())

        # Layout principal
        main_layout = QHBoxLayout()

        # Bouton Accept
        valider_btn = QPushButton("Sauvegarder les modifications")
        valider_btn.clicked.connect(self.valider_btn_action)

        # Bouton Refus
        refus_btn = QPushButton("Modification temporaire")
        refus_btn.clicked.connect(self.refus_btn_action)

        main_layout.addWidget(valider_btn)
        main_layout.addWidget(refus_btn)

        self.setLayout(main_layout)

    def valider_btn_action(self):
        self.sauvegarde.emit(True)
        self.force = True 
        self.close()

    def refus_btn_action(self):
        self.sauvegarde.emit(False)
        self.force = True 
        self.close()

    def closeEvent(self,event):
        if not self.force:
            QApplication.closeAllWindows()
            QApplication.quit()
            event.accept()
            self.fclose.close()

        

# =================================================================================================
# =                                           SECTION: TRAITEMENT 
# =================================================================================================


class Mapping_program_fonction:
    """ Programme principal """

    def __init__(self,file_path,client_name,time,audit,forceclose,absolut_path,Use_info):
        
        
        self.fclose = forceclose                    # Variable de fermeture forcée du programme 
        self.audit = audit                          # Variable de piste d'audit 
        self.file_path = file_path                  # Variable de chemin du fichier d'entrée 
        self.client_name = client_name              # Variable du nom du client 
        self.time = time                            # Instant de la requête
        self.correspondances = {}                   # Dictionnaire de correspondance des colonnes entre celles attendues (clé) et fournies sous forme de liste (valeur) ex : self.correspondance = {'isin' : ['ISIN'] , 'Currency' : ['Cur']}
        self.filling_auto = {}                      # Dictionnaire de remplissage automatique prenant la colonne attendue (clé) et la valeur pour le remplissage (valeur) ex : self.filling_auto {'ValidityType' : 'Day', 'Account' : 'Compte 1'}
        self.notfoundBloom = {}                     # !! Jamais rempli !! ? 
        self.clientUlID = None                      # Variable d'ulID du client 
        self.isin_comparison = {}                   # Dictionnaire contenant les Isins (clé) et les couples ticker_blooms fournis et indices de l'ordre introuvable sous forme de liste (val) ex : self.isin_comparison = {'FR0000000000' : ['AXC MI',3]}
        self.idx_isin = {}                          # Sictionnaire contenant pour chaque Isin introuvable et leur ticker bloom fourni sous forme de liste (val) trouvable à partir de l'indice de l'ordre (clé) ex : self.idx_isin = {3 : ['FR0000000000','AXC MI']}
        self.message = f"Bloomberg Isin et l'Isin indiqué sont différents\n\n" # Variable Message d'erreur  lorsque le code bloom, si il est fourni, ne correspond pas à l'ISIN 
        self.messagecount = False                   # Booléen déterminant si il faut afficher le message d'erreur précédement
        self.Use_info = Use_info                    # Information sur l'utilisateur et le dépôt utilisé (class Useinfo)
        self.absolut_path = absolut_path            # chemin absolut du dossier des programmes

        self.change_memory = False                  # Booléen déterminant la voulonté ou non de modification des informations en mémoires dans la table sql TBL_BTRADE_CLIENTMEMORY du client 
        self.memory = None                          # Liste de dictionnaire de correspondances des colonnes en mémoire du client ou None le client n'est pas encore en mémoire
        self.memory_comparison_client = False       # Booléen indiquant si le format de colonne de fichier du client est déjà en mémoire
        self.rules_data = {}                        # Dictionnaire fournissant les valeurs associé à chaque règle pour le client
        self.col_client_memory_list = []            # Liste des colonnes de la table TBL_BTRADE_CLIENTMEMORY
        self.info_rules = {}                        # Dictionnaire des règles du client en mémoire
        self.isin_info = {}                         # Dictionnaire donnant les informations propre à chaque isin fournis (clé) spécifique à ses tickers self.info = {'FR0000000000' : {'AAA FR' : {"currency": "EUR","exchange": "MTAA","volume": "98925.000000","volume_moyen": "957,084","isin": "IT0004176001","name": "PRYSMIAN SPA","type": "Common Stock"}}}
        self.isin_tickerbloomformat_dic = {}        # Dictionnaire fournissant les isins fournis et les tickers qui n'ont pas correspondu sous forme de liste (val) à partir de l'indice de l'ordre (clé) ex : self.isin_tickerbloomformat_dic = {1 : ['AAA FR', 'FR0000000000']}
        self.ticker_unknown_isin = {}               # Dictionnaire fournissant les informations blooms des tickers introuvables {'AAA FR' : {"currency": "EUR","exchange": "MTAA","volume": "98925.000000","volume_moyen": "957,084","isin": "IT0004176001","name": "PRYSMIAN SPA","type": "Common Stock"}}
        self.exchange_incorrecte = {}               # Dictionnaire des exchanges ne correspondant pas à leurs exchanges blooms fournis 
        self.ticker_unknown_isin_list = []          # Liste des tickers introuvables à partir des isins renseignés 
        self.idx_get_ticker_info_from_tickers = []  # Liste des indices dont les tickers fournis ne fonctionnent pas 
        self.mapping_bloom = {}                     # Dictionnaire de correspondance entre l'operated MIC (val) et le MIC (clé) pour permettre de fournir le MIC général en cas de litige 
        self.etf = {}                               # Dictionnaire des isins fournis etf 
        self.only_bloom_info = {}                   # Dictionnaire d'information des tickers fournis lorsqu'il n'y a pas de colonne d'isin 
        self.no_isin = False                        # Booleen indiquant la présence ou non d'une colonne d'isin 
        self.message_no_isin = ''
        
        
        
        
        self.nkey = ['isin','Bloomberg Code','Exchange','Currency','Side','Quantity','OrderType','ValidityType','Account','Price','StopPrice','ExpireDate','Reference','Comment', 'ValidityType & OrderType']
        self.Lkey = ['isin','Bloomberg Code','Exchange','Currency','Quantity','Price','StopPrice','ExpireDate','Reference','Comment']
        self.Ckey =  ['Side','ValidityType','OrderType','Account']

        
        # Optimisation: Le nettoyage et filtrage sont  intégrés dans convert_file()
       
        self.csv_path = self.convert_file()
    
        

       
        self.df = pd.read_csv(self.csv_path)
        
        

       
        self.col_client_memory()


       
        self.determineClientUlId()
 
        
       
        self.client_accounts, self.accounts_dic =  get_client_accounts(self.clientUlID)


        self.audit.info(f"Comptes du client : {self.client_accounts}")
        self.audit.info(f"Informations des comptes : {self.accounts_dic}")

       
        self.Client_memory()


                 
        # Creation des differents dictionnaires de mappings pour les differentes colonnes
        self.mapping_bloom = self.mappingbloom()
        
        self.mapping_currency = self.mappingcurrency_constructor()
        self.mapping_side = self.mappingside_constructor()
        self.mapping_type = self.mappingtype_constructor()
        self.mapping_validitytype = self.mappingvaliditytype_constructor()
        #self.etf = self.mappingetf_constructor()
       
        self.verif_window()
   

        self.add_isin()
            
            
        self.formatage_isin()
       
        self.change_database_client_memory()


       
        self.sauvegarde()



       
        self.isin_bloom_info()
 

       
        self.complete_bloom_columns()

        

       
        if self.ticker_unknown_isin_list : 
            input = {'tickers' : self.ticker_unknown_isin_list,'idx_list' : self.idx_get_ticker_info_from_tickers}
            self.ticker_unknown_isin = discussion('ticker_info',input,self.Use_info.user_id)
       

       
        self.msg_isin_bloom()


        # Remplie des pandas Series pour contruire l'excel d'ordre 
       
        self.ISIN = self.verif_isin()


       
        self.EXCHANGE = self.verif_exchange()
      

       
        self.CURRENCY = self.verif_currency()

       
        self.SIDE = self.verif_side()
  
       
        self.QUANTITY = self.verif_quantity()
      

       
        self.ORDERTYPE = self.verif_type()
        

       
        self.VALIDITYTYPE = self.verif_validitytype()

       
        self.ACCOUNT = self.verif_account()
   


       
        self.PRICE = self.verif_price()
        

       
        self.STOPPRICE = self.verif_stopprice()
   

       
        self.EXPIREDATE = self.verif_expiredate()
      
       
        self.REFERENCE = self.verif_reference()
  
       
        self.COMMENT = self.verif_comment()


       
        self.rules()
       
        self.affichage_notfoundBloom()

        
    #---------------------------------------------------------------------------------    

    def add_isin(self):
        "Fonction de création de la colonne d'ISIN lorsqu'elle n'existe pas"
        if 'isin' not in self.correspondances:
            self.no_isin = True
            bloom_col = self.correspondances['Bloomberg Code'][0]
            bloom_code_liste = self.df[bloom_col].apply(self.extract_bloom)
            bloom_code_liste = bloom_code_liste.apply(bloom_format_transformation_ticker)
            idx_list = []
            for i in range(len(bloom_code_liste)):
                idx_list.append(i)

            input = {'tickers' : bloom_code_liste.tolist(),'idx_list' : idx_list}
            list_result = discussion('ticker_info',input,self.Use_info.user_id)
            
           
            self.df['ISIN'] = None 
            self.correspondances['isin'] = ['ISIN']
            message = ''
            for code, info in list_result.items():
                self.df.iloc[info["idx"],self.df.columns.get_loc("ISIN")] = info['isin']
                if info['currency'] == 'Not Valid':
                    message += f"- Ordre : {info["idx"]} Ticker {code} Introuvable \n"

            self.only_bloom_info = list_result
            if message != '':
                self.message_no_isin = message
                
    #---------------------------------------------------------------------------------

    def isin_bloom_info(self):
        "Fonction de récupération des données des isin à l'aide du canal de discussion avec Bloom"
        if 'isin' not in self.correspondances and self.no_isin == False:
            self.audit.error(f"Aucune colonne correspondant à l'ISIN ")
            raise ValueError("Aucune colonne correspondant à l'ISIN ")
        if self.no_isin == False:
            col_isin = self.correspondances ['isin']
            col_isin = col_isin[0]

            isin_list = self.df[col_isin].tolist()
            self.isin_info = discussion(input = isin_list, fonction = 'isin_info',user_id = self.Use_info.user_id)

    def formatage_isin(self):
        "Formatage des Isins mals renseignés"
        if 'isin' not in self.correspondances and self.no_isin ==False:
            raise(KeyError,'Error isin')
            
        isin_col = self.correspondances['isin'][0]
        self.df[isin_col] = self.df[isin_col].str.replace(r"-.*","",regex = True)

    #---------------------------------------------------------------------------------

    def verif_window(self):
        """Ouvre la fenêtre de vérification des colonnes et retourne les informations avec transformation"""
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        conn = pymssql.connect(
        server = '',
        user = "",
        password = '',
        database = ''
    ) 
            
        cursor = conn.cursor(as_dict=True)
        query = "SELECT Case_unique_account, unique_currency, unique_exchange, supress_order_on_quantity, maximum_volume_exchange FROM TBL_BTRADE_CLIENTMEMORY WHERE ClientName = %s  "
        params = [self.clientUlID]

        cursor.execute(query,params)
        row = cursor.fetchall()
        if row :
            self.info_rules = row[0]
        
        
        self.verif_window = DataFrameWidget(self.fclose,self.clientUlID, self.df, self.nkey,self.Lkey,self.Ckey,get_value_for_key, get_default_key_for_column,self.client_accounts, self.memory,self.memory_comparison_client, self.info_rules,self.mapping_currency,self.mapping_side,self.mapping_type,self.mapping_validitytype)
        self.verif_window.validated.connect(self.transformation)
        self.verif_window.show()
        app.exec()

    def transformation(self,dic_correspondance,dic_auto, dic_rules):
        """Alloue :
                - self.correspondance un dictionnaire prenant en clé les colonnes attendues et en valeur la colonne correspondante
                - self.filling_auto un dictionnaire prenant en clé la colonne à remplir et en valeur la valeur à remplir dans l'ensemble des lignes
                - self.dic_rules un dictionnaire de règles du client"""
        self.correspondances = dic_correspondance
        self.filling_auto = dic_auto
        self.dic_rules = dic_rules
        
    #---------------------------------------------------------------------------------

    def determineClientUlId(self):
        """ Alloue à self.ClientUlId l'UlId du client à partir de la table TBLCLIENTS
            Si UlId introuvable, passe par une demande directe 
                - SelectionWindow si plusieurs résultats possibles
                - ClientIdInputWindow si aucun resultat trouvé"""
        
        conn = pymssql.connect(
                server = '',
                user = '',
                password= '',
                database=''
            )
        cursor = conn.cursor()
        cursor.execute("SELECT Ul_Id FROM TBLCLIENTS WHERE Description = %s", (self.client_name,))
        result = cursor.fetchall()
        conn.close()

        self.audit.info(f"Client trouvée : {result}")

        if result:
            liste = [item[0] for item in result]
            descriptions = set(liste)
            if len(descriptions) == 1:
                self.clientUlID = descriptions.pop()
            else:
                self.clientIdWindow = SelectionWindow(descriptions,self.fclose)
                # Connecter le signal pour recevoir les mises à jour
                app = QApplication.instance()
                if app is None:
                    app = QApplication(sys.argv)
                self.clientIdWindow.selection_ready.connect(self.apply_selected_Ul_Id)
                self.clientIdWindow.show()
                app.exec()
        else:
            self.clientIdWindow = ClientIdInputWindow(self.fclose)
            # Connecter le signal pour recevoir les mises à jour
            app = QApplication.instance()
            if app is None:
                app = QApplication(sys.argv)
            self.clientIdWindow.client_id_ready.connect(self.apply_selected_Ul_Id)
            self.clientIdWindow.show()
            app.exec()

    def apply_selected_Ul_Id(self, selected_id):
        """Appliquer l'Ul_Id sélectionné"""
        self.clientUlID = selected_id

        self.audit.info(f"Ul_Id du client : {self.clientUlID}")

    #----------------------------------------------------------------------------------

    def Client_memory(self):
        """ Alloue :
                - self.columns_memory la liste des colonnes de TBL_BTRADE_CLIENTMEMORY
                - self.rules_data le dictionnaire des règles spécifiques du client
                - self.memory la liste de dictionnaire de correspondances des colonnes en mémoire du client ou None si une colonne ne correspond pas 
                - self.memory_comparison_client"""
                
        conn = pymssql.connect(
        server = '',
        user = "",
        password = '',
        database = ''
        )
        cursor = conn.cursor(as_dict= True)
        query = f"SELECT  * FROM TBL_BTRADE_CLIENTMEMORY WHERE ClientName = %s "
        params = self.clientUlID
        
        cursor.execute(query,params)
        rows = cursor.fetchall()
        query2  = "SELECT TOP 1 * FROM TBL_BTRADE_CLIENTMEMORY "

        cursor.execute(query2)
        self.columns_memory = list(cursor.fetchall()[0].keys())

        self.rules_data = {rules_col : None for rules_col in  self.columns_memory[18:]}

        if rows != []:
            memory_column_name = rows[0]
            self.memory = memory_column_name
            self.memory_comparison_client = self.memory_comparison()

            self.audit.info(f"Informations en memoire : {self.memory}")

            if self.memory_comparison_client:
                self.audit.info(f"Correspondance Fichier/Memoire : {self.memory_comparison_client}")
            else:
                self.audit.warning(f"Correspondance Fichier/Memoire : {self.memory_comparison_client}")

        else:
            self.memory = None

            self.audit.info(f"Client abscent de la memoire")

    def memory_comparison(self):
        """Compare les colonnes du fichier avec ceux en mémoire et retourne un booleen"""

        df_columns = [col.upper() for col in self.df.columns]
        
        list_values = list(self.memory.values()) 
        list_col = list_values[1:14] + [list_values[17]]
        
        for elt in list_col:
            if elt:
                
                if elt.upper() not in df_columns:
                    self.audit.warning(f"Colonne {elt} Introuvable")
                    
                    return False
        # separation des multicolonnes de Comment
        if list_values[14]:
            list_comment_col = list_values[14].split('#')
            for elt in list_comment_col:
                if elt:
                    if elt.upper() not in df_columns:
                        self.audit.warning(f"Colonne {elt} Introuvable")
                        return False 
        
        self.rules_data = {rules_col : self.memory[rules_col] for rules_col in  list(self.memory.keys())[18:]}
        return True
                    
    def change_database_client_memory(self):
        """Propose de mettre en mémoire les modification ou non"""
        self.change_memory = False

        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)
        window = Demande_mise_en_memoire()
        window.sauvegarde.connect(self.sauvegarde_modification)
        window.show()
        app.exec()
    
    def sauvegarde_modification(self,save):
        """ Modifie le booleen de mise en mémoire ou non des modifications """
        self.audit.info(f"Mise en memoire : {save}")
        self.change_memory = save
    
    def sauvegarde(self):
        """Si la self.change_memory == True:
                - Suppression les inforamtions du clien en mémoir de TBL_BTRADE_CLIENTMEMORY
                - Ajout des nouvelles informations"""
        
        if self.change_memory:
            self.conn =  pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
            self.cursor = self.conn.cursor()

            query_sup = " DELETE FROM TBL_BTRADE_CLIENTMEMORY WHERE ClientName = %s "
            params_sup = self.clientUlID
            self.cursor.execute(query_sup,params_sup)
            self.conn.commit()

            col_str = ', '.join([f"[{c}]" for c in self.col_client_memory_list[:9]]) + ', Date,' + ', '.join([f"[{c}]" for c in self.col_client_memory_list[9:19]]) + ', ' + ', '.join([f"[{c}]" for c in self.col_client_memory_list[20:]])
            val_str = ','.join(['%s'] * len(self.col_client_memory_list))
            query = f'''INSERT INTO TBL_BTRADE_CLIENTMEMORY
                        ({col_str})
                        VALUES ({val_str})
                    '''
            
            list_data_save = []
            if 'ValidityType' in self.correspondances:
                col_validityType = self.correspondances['ValidityType'][0]
                col_date = None
            elif 'Date' in self.correspondances:
                col_date = self.correspondances['Date'][0]
                col_validityType = None

            col_orderType = self.correspondances['OrderType'][0]

            if col_validityType == 'Manuel' and 'ValidityType' in list(self.filling_auto.keys()):
                list_data_save.append(self.filling_auto['ValidityType'])
            else:
                list_data_save.append(None)
            
            if col_orderType == 'Manuel' and 'OrderType' in list(self.filling_auto.keys()):
                list_data_save.append(self.filling_auto['OrderType'])
            else:
                list_data_save.append(None)
            
            
            order_list_1 = ['Exchange','Currency','Bloomberg Code','Side','Quantity']
            order_list_int =['OrderType','ValidityType','Date']
            order_list_2 = ['Price','StopPrice','ExpireDate','Account','Reference','Comment']

            list_col_save_1 = [None if self.correspondances.get(cle) == ['Manuel'] else self.correspondances.get(cle,None) for cle in order_list_1]
            list_col_save_2 = [None if self.correspondances.get(cle) == ['Manuel'] else self.correspondances.get(cle,None) for cle in order_list_2]

            if not self.no_isin:
                list_isin = [None if self.correspondances.get(cle) == ['Manuel'] else self.correspondances.get(cle,None) for cle in ['isin']]
            else:
                list_isin = [None]
            list_col_save_1 = list_isin + list_col_save_1

            


            if isinstance(list_col_save_2[-1],list) and len(list_col_save_2[-1]) >1 :
                list_col_save_2[-1] = ['#'.join(list_col_save_2[-1])]

            list_validity_order = [self.correspondances.get('ValidityType & OrderType',None)]

            if 'ValidityType & OrderType' in self.correspondances:
                list_col_save_int = [None,None,None]
            else:
                list_col_save_int = [None if self.correspondances.get(cle) == ['Manuel'] else self.correspondances.get(cle,None) for cle in order_list_int]

            list_rules = []
            for rule in self.dic_rules.keys():
                if self.dic_rules[rule]['checkbox']:
                    text = self.dic_rules[rule]['edit'] 
                    if text :
                        list_rules.append(text)
                    else:
                        list_rules.append(None)
                else:
                    list_rules.append(None)

            params = tuple([self.clientUlID] + list_col_save_1 + list_col_save_int + list_col_save_2 + list_data_save + list_validity_order + list_rules )
            
            self.audit.info(f"Client {self.clientUlID}, Memorisation : {params}")
            
            self.cursor.execute(query,params)
            self.conn.commit()
            self.conn.close()
            return 

    #---------------------------------------------------------+----------------------------------------------------------------------------------------------------------------------    
    
    def convert_file(self):

        """La fonction étudie le fichier fourni : 
                - CSV ou XLSX
                - Titre de colonne ou non 
                - Colonnes vides
            Puis crée un nouveau fichier au format permettant l'analyse"""
        self.format = ''
        _, ext = os.path.splitext(self.file_path)
        ext = ext.lower() 
        
        folder = os.path.dirname(self.file_path) or '.'
        csv_filename = f"{self.client_name}-{self.time}.csv"
        csv_path = os.path.join(folder, csv_filename)
        
        
        try:
            if ext == ".xlsx":
                self.format = 'xlsx'
                # Détection optimisée de la ligne d'en-tête pour Excel

          
                
                # Méthode vectorisée pour détecter l'en-tête
                
                
                df_raw = pd.read_excel(self.file_path, header=None, nrows=10)
                df_raw = df_raw.replace(r'^\s*$', np.nan, regex = True)
                
                pd.set_option("display.max_rows",None)
                pd.set_option("display.max_columns",None)
                
                pd.set_option("display.max_colwidth",None)
               
                
                non_null_counts = df_raw.notna().sum(axis=1)
                
                header_row = non_null_counts[non_null_counts >= 5].index[0] if (non_null_counts >= 5).any() else 0
              
                
                
                
                df_int = pd.read_excel(self.file_path,header=header_row)
                df_int = df_int.replace(r'^\s*$', np.nan, regex = True)
                df_int = df_int.iloc[:, :20]
                

                
                has_float = any(self.is_float(cell) for cell in df_int.columns)
                
                
                if has_float:
                    first_row = pd.DataFrame([df_int.columns], columns=df_int.columns)
                    
                    df_int = pd.concat([first_row,df_int],ignore_index=True)
                    
                    df_int.columns = [f'Column_{i}'for i in range(df_int.shape[1])]
                    
                    df = df_int
                    
                
                else:
                    
                    df = df_int


            elif ext == ".csv":
                self.format = 'csv'
                with open(self.file_path, "r", encoding="utf-8") as f:
                    first_line = f.readline().strip().split(';')
                    
                
                has_float = any(self.is_float(cell) for cell in first_line)

                if has_float:
                    df_raw = pd.read_csv(self.file_path, header=None, nrows=10, sep=';')
                    df_raw.columns = [f'Column_{i}'for i in range(df_raw.shape[1])]
                else:
                    df_raw = pd.read_csv(self.file_path, nrows=10, sep=';')
                
                
                # Méthode vectorisée pour détecter l'en-tête
                non_null_counts = df_raw.notna().sum(axis=1)
                header_row = non_null_counts[non_null_counts >= 3].index[0] if (non_null_counts >= 3).any() else 0
                
                
                
                # Lecture unique avec la bonne ligne d'en-tête
                df = pd.read_csv(self.file_path, header=header_row, sep = ';')

                if has_float:
                    df.columns = [f'Column_{i}'for i in range(df.shape[1])]
                
                
            else:
                raise TypeError("Only csv or xlsx file")
            
            # Nettoyage optimisé du DataFrame avec opérations vectorisées
            # Calcul vectorisé du nombre d'éléments valides par ligne
            df_str = df.astype(str)
            mask_not_empty = (df_str != '') & (df_str != 'nan')
            valid_count_per_row = mask_not_empty.sum(axis=1)
            
            # Filtrage vectorisé des lignes avec moins de 4 éléments valides
            df_cleaned = df[valid_count_per_row >= 4].reset_index(drop=True)
            
            # Nettoyage des noms de colonnes intégré ici pour éviter une opération supplémentaire
            df_cleaned.columns = df_cleaned.columns.str.replace(r'[\n\r\t]', '', regex=True).str.strip()
            
            # Sauvegarder en CSV
            df_cleaned.to_csv(csv_path, index=False)
            
            return csv_path
            
        except Exception as e:
            print(f"Erreur lors de la conversion : {e}")
            return None
 
    def excel_extract_value(self,pos):
        """Extrait l'élément en position pos dans le fichier d'ordre initial"""

        if self.format == "xlsx":
            df_lecture = pd.read_excel(self.file_path)
        if self.format == "csv":
            df_lecture = pd.read_csv(self.file_path)
        val = excel_to_df(df_lecture,pos)
        return val
            
    #-----------------------------------------

    # FONCTIONS DE MAPPING A PARTIR DES TABLES SQL

    def mappingbloom(self):
        """Mappping MIC et OPERATED MIC"""
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT exchange_bloom, exchange_ul FROM TBL_BTRADE_BLOOMEXCHANGE")

        mapping = {}
        for bloom, ul in cursor.fetchall():
                mapping[bloom] = ul

        conn.close()
        cursor.close()
        return mapping

    def mappingcolumn_constructor(self):
        """Mapping colonnes attendus et leurs variantes pour identifier les colonnes """
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT colonne_attendue, colonne_variante FROM TBL_BTRADE_COLNAME")

        mapping = {}
        for attendu, variante in cursor.fetchall():
            if attendu in self.selected_columns_auto:
                mapping.setdefault(attendu,[]).append(variante)

        conn.close()
        cursor.close()
        return mapping
     
    
    def mappingcurrency_constructor(self):
        """Mapping de correspondance des currencies"""
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT currency_attendue, currency_variante FROM TBL_BTRADE_CURRENCY")

        mapping = {}
        for attendu, variante in cursor.fetchall():
            mapping.setdefault(attendu,[]).append(variante)

        conn.close()
        cursor.close()
        return mapping

    def mappingside_constructor(self):
        "Mapping de correspondance des sens de l'ordre "
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT side_attendue, side_variante FROM TBL_BTRADE_SIDE")

        mapping = {}
        for attendu, variante in cursor.fetchall():
            mapping.setdefault(attendu,[]).append(variante)

        conn.close()
        cursor.close()
        return mapping

    def mappingtype_constructor(self):
        "Mapping de Ordertype"
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT type_attendue, type_variante FROM TBL_BTRADE_TYPE")

        mapping = {}
        for attendu, variante in cursor.fetchall():
            mapping.setdefault(attendu,[]).append(variante)

        conn.close()
        cursor.close()
        return mapping

    def mappingvaliditytype_constructor(self):
        """Mapping de Validitytype"""
        conn = pymssql.connect(
                        server = '',
                        user = "",
                        password = '',
                        database = ''
                        )
        cursor = conn.cursor()
        cursor.execute("SELECT validitytype_attendue, validitytype_variante FROM TBL_BTRADE_VALIDITYTYPE")

        mapping = {}

        for attendu, variante in cursor.fetchall():
            mapping.setdefault(attendu,[]).append(variante)

        conn.close()
        cursor.close()
        return mapping


    def verif_isin(self):
        """Fonction de vérification des isins retournant une pandas Series d'isin"""
        if self.no_isin:
            # Absence d'isins fournis dans le fichier donc aucune comparaison nécessaire
            list_isin = self.df['ISIN'].tolist()
            ISIN_code = pd.Series(list_isin, index = self.df.index, name = 'ISIN_Code')
            return ISIN_code
        else :
            
            if 'isin' not in self.correspondances:
                self.audit.error(f"Aucune colonne correspondant à l'ISIN ")
                raise ValueError("Aucune colonne correspondant à l'ISIN ")
            
            col_isin = self.correspondances['isin']
            col_isin = col_isin[0]
            
            isin_list = []

            # Ajout pour colonne manuelle
            if col_isin == "Manuel":
                if 'isin' in list(self.filling_auto.keys()):
                    # Ajout d'un isin pour tout les ordres (rare)
                    for idx in self.df.index:
                        isin_list.append(self.filling_auto['isin'])
                        self.audit.info(f" ISIN, Ordre : {idx}, col : {col_isin}, code : {self.filling_auto['isin']} ")
                else:
                    for idx in self.df.index:
                        isin_list.append('Not Found')
                        self.audit.warning(f" ISIN Abscent (manuel), Ordre : {idx}")
                    ISIN_code = pd.Series(isin_list, index = self.df.index, name = 'ISIN_Code')
                return ISIN_code

            
            for idx, val in self.df[col_isin].items():
                if val == '' or pd.isna(val):
                    isin_list.append('Not Found')
                    self.audit.warning(f" Isin Abscent, Ordre : {idx}, colonne : {col_isin}")
                else:
                    val = self.extract_isin(val)
                    if val :
                            # Isin toujours indiqué même si celui ci est faux. Seul l'absence d'isin est indiqué ici (les erreurs sont dans la colonne exchange ou dans la couleur d'affichage)
                            isin_list.append(val.strip())
                            self.audit.info(f" ISIN, Ordre : {idx}, col : {col_isin}, code : {val}, Trouvé ")
                    else:
                        self.audit.warning(f" ISIN, Ordre {idx} dans la colonne {col_isin} : {val} , Invalide ")
                        isin_list.append('Not Valid')

            ISIN_code = pd.Series(isin_list, index = self.df.index, name = 'ISIN_Code')
            return ISIN_code


    def verif_exchange(self):
        """ Verification de l'exchange fourni et celui attendu  """
        if self.no_isin:
            list_exchange = self.df['Exchange'].tolist()
            
            MIC_code = pd.Series(list_exchange, index = self.df.index, name = 'MIC_Code')
            return MIC_code
        else :
        
            mic_list = []
            
            if 'Exchange' not in self.correspondances:
                self.audit.warning(f"Exchange Abscent , Nouvelle colonne : Not Found ") 
                for idx in range(len(self.df)):
                    mic_list.append('Not Found')
            else:         
                col_exchange = self.correspondances ['Exchange']
                col_exchange = col_exchange[0]
            
                if col_exchange == '':
                    for idx in self.df.index:
                        mic_list.append('Not Found')
                        self.audit.warning(f" Exchange Abscent, Ordre : {idx}, colonne : {col_exchange}")
                    MIC_code = pd.Series(mic_list, index = self.df.index, name = 'MIC_Code')
                    return MIC_code

                

                # Ajout pour colonne manuelle
                if col_exchange == "Manuel":
                    if 'Exchange' in list(self.filling_auto.keys()):
                        self.audit.info(f" Exchange, col : {col_exchange}, code : {self.filling_auto['Exchange']} ")
                        for idx in self.df.index:
                            mic_list.append(self.filling_auto['Exchange'])
                            
                    else:
                        self.audit.warning(f" Exchange Abscent (manuel)")
                        for idx in self.df.index:
                            mic_list.append('Not Found')
                            
                    MIC_code = pd.Series(mic_list, index = self.df.index, name = 'MIC_Code')
                    return MIC_code

                if 'Bloomberg Code' in self.correspondances:
                    for idx,ticker in enumerate(self.df[self.correspondances['Bloomberg Code'][0]]):
                        ticker = self.extract_bloom(ticker)
                        ticker =  bloom_format_transformation_ticker(ticker)
                        col_isin = self.correspondances['isin'][0]
                        exchange = self.df[self.correspondances['Exchange'][0]].iloc[idx]
                        if exchange == '':
                            if ticker and (ticker not in self.ticker_unknown_isin)  :
                                isin = self.df[col_isin].iloc[idx]
                                if self.isin_info[isin] :
                                    
                                    d = self.isin_info[self.df[col_isin].iloc[idx]][ticker]
                                    if d :
                                        exchange_bloom = d['exchange']
                                        
                                        if exchange_bloom in self.mapping_bloom:
                                            exchange_bloom = self.mapping_bloom[exchange_bloom]
                                            
                                        mic_list.append(exchange_bloom)
                                    else:
                                        mic_list.append('Not Valid')
                                else:
                                    mic_list.append('Not Valid')
                            else : 
                                mic_list.append('Not Valid')
                        else:
                            
                            if ticker and ticker not in self.ticker_unknown_isin :
                                isin = self.df[col_isin].iloc[idx]
                                if self.isin_info[isin]:
                                    d = self.isin_info[self.df[col_isin].iloc[idx]][ticker]
                                    if d :
                                        exchange_bloom = d['exchange']
                                        
                                        if exchange_bloom in self.mapping_bloom:
                                            exchange_bloom = self.mapping_bloom[exchange_bloom]
                                        if exchange == exchange_bloom:
                                            mic_list.append(exchange)
                                        else : 
                                            mic_list.append(exchange_bloom)
                                            self.audit.warning(f"Ordre : {idx}, exchange attendu {exchange}, exchange Bloom = {exchange_bloom}")
                                            self.exchange_incorrecte[idx] = {'fourni' : exchange, 'attendu' : exchange_bloom}
                                    else:
                                        mic_list.append('Not Valid')
                                else: 
                                    mic_list.append('Not Valid')
                            else:
                                mic_list.append('Not Valid')

                    MIC_code = pd.Series(mic_list, index = self.df.index, name = 'MIC_Code')
                    return MIC_code





                conn = pymssql.connect(
                        server = '',
                        user = '',
                        password= '',
                        database=''
                    )
                cursor = conn.cursor(as_dict = True)

                try:
                    cursor.execute("SELECT Acronym, Description, Ul_Id FROM TBLEXCHANGES")
                    rows = cursor.fetchall()
                    ref = pd.DataFrame(rows)
                finally:
                    conn.close()

                def _norm(x):
                    if pd.isna(x):
                        return None
                    return str(x).strip().casefold()

                mapping = {}

                for _, row in ref.iterrows():
                    ul = row.get('Ul_Id')
                    for col in ('Acronym','Description','Ul_Id'):
                        raw = row.get(col)
                        k = _norm(raw)
                        if k:
                            if k not in mapping :
                                mapping[k] = ul

                for idx, raw_val in self.df[col_exchange].items():
                    if pd.isna(raw_val) or str(raw_val).strip() =='':
                        self.audit.warning(f" Exchange Abscent, Ordre : {idx}, colonne : {col_exchange}")
                        mic_list.append('Not Found')
                        continue
                    key = _norm(raw_val)
                    
                    ul_found = mapping.get(key)
                    

                    if ul_found is not None :
                        self.audit.info(f" Exchange, Ordre : {idx}, col : {col_exchange}, code : {ul_found}, Trouvé ")
                        if not self.no_isin : 
                            col_isin = self.correspondances['isin'][0]
                            isin = self.df[col_isin].iloc[idx]
                            if self.isin_info[isin] == None:
                                self.audit.warning(f" Ex - Isin, Ordre {idx} dans la colonne {col_isin} : {isin} , Invalide ")
                                mic_list.append('ISIN Invalide')
                            else:
                                mic_list.append(ul_found)
                        else:
                            mic_list.append(ul_found)
                    else:

                        self.audit.warning(f" Exchange, Ordre {idx} dans la colonne {col_exchange} : {raw_val} , Invalide ")
                        mic_list.append('Not Valid')

            MIC_code = pd.Series(mic_list, index = self.df.index, name = 'MIC_Code')
            return MIC_code

    def verif_currency(self):
        """ Verification de la currency """
        if self.no_isin:

            list_currency= self.df['Currency'].tolist()
            
            CURRENCY_code = pd.Series(list_currency, index = self.df.index, name ='CURRENCY_Code')
            return CURRENCY_code

        currency_list = []
        if 'Currency' not in self.correspondances:
            self.audit.warning(f"Currency Abscent, Nouvelle colonne ") 
            for idx in range(len(self.df)):
                currency_list.append('Not Found')
        else:
            col_currency = self.correspondances ['Currency']
            col_currency = col_currency[0]

            currency_list = []


            if col_currency == '':
                self.audit.warning(f" Currency Abscent, colonne : {col_currency}")
                for idx in self.df.index:
                    currency_list.append('Not Found')
                   
                CURRENCY_code = pd.Series(currency_list, index = self.df.index, name ='CURRENCY_Code')
                return CURRENCY_code
            
            # Ajout pour colonne manuelle
            if col_currency == "Manuel":
                self.audit.info(f" Currency, col : {col_currency}, code : {self.filling_auto['Currency']} ")
                if 'Currency' in list(self.filling_auto.keys()):
                    for idx in self.df.index:
                        currency_list.append(self.filling_auto['Currency'])
                        
                else:
                    for idx in self.df.index:
                        currency_list.append('Not Found')
                        self.audit.warning(f" Currency Abscent (manuel), Ordre : {idx}")
                CURRENCY_code = pd.Series(currency_list, index = self.df.index, name ='CURRENCY_Code')
                return CURRENCY_code

            

            for idx, val  in self.df[col_currency].items():
                if val == '' or pd.isna(val):
                    self.audit.warning(f" Currency Abscent, Ordre : {idx}, colonne : {col_currency}")
                    currency_list.append('Not Found')
                else:
                    find = None
                    for key, variante  in self.mapping_currency.items():
                        variante = [x.upper() for x in variante]
                        val = val.upper()
                        if val in variante :
                            find = key
                            break
                    if find :
                        self.audit.info(f" Currency, Ordre : {idx}, col : {col_currency}, code : {val}, Trouvé ")
                        currency_list.append(find)
                    else:
                        self.audit.warning(f" Currency, Ordre {idx} dans la colonne {col_currency} : {val} , Invalide ")
                        currency_list.append('Not Valid')

        CURRENCY_code = pd.Series(currency_list, index = self.df.index, name ='CURRENCY_Code')
        return CURRENCY_code

    def verif_side(self):
        """Verification du side à l'aide de la table de correspondance TBL_BTRADE_CURRENCY"""
        if 'Side' not in self.correspondances:
            self.audit.error(f"Aucune colonne correspondant au Side ")
            raise ValueError("Aucune colonne correspondant à side ")
        col_side = self.correspondances ['Side']
        col_side = col_side[0]
        side_list = []

        # Ajout pour colonne manuelle
        if col_side == "Manuel":
            if 'Side' in list(self.filling_auto.keys()):
                self.audit.info(f" Side, col : {col_side}, code : {self.filling_auto['Side']} ")
                for idx in self.df.index:
                    side_list.append(self.filling_auto['Side'])
                    
            else:
                self.audit.warning(f" Side Abscent (manuel)")
                for idx in self.df.index:
                    side_list.append('Not Found')
                    
            SIDE_code = pd.Series(side_list, index = self.df.index, name ='SIDE_Code')
            return SIDE_code

        
        for idx, val  in self.df[col_side].items():
            if val == '' or pd.isna(val):
                self.audit.warning(f" Side Abscent, Ordre : {idx}, colonne : {col_side}")
                side_list.append('Not Found')
            else:
                
                find = None
                for key, variante  in self.mapping_side.items():
                    variante = [x.upper().strip() for x in variante]
                    val = val.upper().strip()
                    if val in variante :
                        find = key
                        break
                if find :
                    self.audit.info(f" Side, Ordre : {idx}, col : {col_side}, code : {val}, Trouvé ")
                    side_list.append(find)
                else:
                    self.audit.warning(f" ISIN, Ordre {idx} dans la colonne {col_side} : {val} , Invalide ")
                    side_list.append('Not Valid')

        SIDE_code = pd.Series(side_list, index = self.df.index, name ='SIDE_Code')
        return SIDE_code

    def verif_quantity(self):
        """Verification du format de la quantité et passage en entier naturel"""
        if 'Quantity' not in self.correspondances:
            raise ValueError("Aucune colonne correspondant à quantity ")
        col_quantity = self.correspondances ['Quantity']
        col_quantity = col_quantity[0]
        quantity_list = []

        # Ajout pour colonne manuelle
        if col_quantity == "Manuel":
            self.audit.info(f" Quantity, col : {col_quantity}, code : {self.filling_auto['Quantity']} ")
            if 'Quantity' in list(self.filling_auto.keys()):
                for idx in self.df.index:
                    quantity_list.append(self.filling_auto['Quantity'])
                    
            else:
                self.audit.warning(f" Quantity Abscent (manuel)")
                for idx in self.df.index:
                    quantity_list.append('Not Found')
                    
            QUANTITY_code = pd.Series(quantity_list, index = self.df.index, name = 'QUANTITY_Code', dtype=object)
            return QUANTITY_code


        for idx, val in self.df[col_quantity].items():
            if val == '' or pd.isna(val):
                self.audit.warning(f" Quantity Abscent, Ordre : {idx}, colonne : {col_quantity}")
                quantity_list.append('Not Found')
            else:
                val_int = self.to_int(val)
                

                if val_int is not None:
                    self.audit.info(f" Quantity, Ordre : {idx}, col : {col_quantity}, code : {val_int}, Trouvé ")
                    quantity_list.append(val_int)
                else:
                    self.audit.warning(f" ISIN, Ordre {idx} dans la colonne {col_quantity} : {val} , Invalide ")
                    quantity_list.append('Not Valid')

        QUANTITY_code = pd.Series(quantity_list, index = self.df.index, name = 'QUANTITY_Code', dtype=object)
        return QUANTITY_code

    def verif_type(self):
        """Verification de l'Ordertype"""
        if 'OrderType' not in self.correspondances:
            
            self.audit.error(f"Aucune colonne correspondant à l'OrderType ")
            raise ValueError("Aucune colonne correspondant à ordertype ")
        col_type = self.correspondances['OrderType']
        col_type = col_type[0]
        type_list = []

        # Ajout pour colonne manuelle
        if col_type == "Manuel":
            if 'OrderType' in list(self.filling_auto.keys()):
                self.audit.info(f" OrderType, col : {col_type}, code : {self.filling_auto['OrderType']} ")
                for idx in self.df.index:
                    type_list.append(self.filling_auto['OrderType'])
                    
            else:
                self.audit.warning(f" OrderType Abscent (manuel)")
                for idx in self.df.index:
                    type_list.append('Not Found')
                    
            TYPE_code = pd.Series(type_list, index=self.df.index, name='ORDERPRICE_Code')
            return TYPE_code

        for idx, val  in self.df[col_type].items():
            if val == '' or pd.isna(val):
                self.audit.warning(f" OrderType Abscent, Ordre : {idx}, colonne : {col_type}")
                type_list.append('Not Found')
            else:
                find = None
                for key, variante  in self.mapping_type.items():
                    variante = [x.upper().strip() for x in variante]
                    val = val.upper()
                    val = val.strip()
                    
                    if val in variante :
                        find = key
                        break
                if find :
                    self.audit.info(f" OrderType, Ordre : {idx}, col : {col_type}, code : {val}, Trouvé ")
                    type_list.append(find)
                else:
                    self.audit.warning(f" OrderType, Ordre {idx} dans la colonne {col_type} : {val} , Invalide ")
                    type_list.append('Not Valid')

        TYPE_code = pd.Series(type_list, index = self.df.index, name ='ORDERPRICE_Code')
        return TYPE_code

    def verif_validitytype(self):
        """Verification du ValidityType"""
        if 'ValidityType' not in self.correspondances and 'Date' not in self.correspondances:
            self.audit.error(f"Aucune colonne correspondant à Validitytype ")
            raise ValueError("Aucune colonne correspondant à ordertype ")
        

        elif 'ValidityType' not in self.correspondances and 'Date' in self.correspondances:
            validitytype_list = []
            col_date = self.correspondances['Date'][0]
           
            for idx, val in self.df[col_date].items():
               
                date = date_transform(val)
                validitytype_list.append(date)
                self.audit.info(f"Validitytype par Date : Ordre = {val}, val = {date}")
            VALIDITYTYPE_code = pd.Series(validitytype_list, index = self.df.index, name ='VALIDITYTYPE_Code')
            return VALIDITYTYPE_code

            
        col_validitytype = self.correspondances ['ValidityType']
        col_validitytype = col_validitytype[0]
        

        validitytype_list = []
        
        # Ajout pour colonne manuelle
        if col_validitytype == "Manuel":
            if 'ValidityType' in list(self.filling_auto.keys()):
                self.audit.info(f" ValidityType, col : {col_validitytype}, code : {self.filling_auto['ValidityType']} ")
                for idx in self.df.index:
                    validitytype_list.append(self.filling_auto['ValidityType'])
                   
            else:
                self.audit.warning(f" ValidityType Abscent (manuel)")
                for idx in self.df.index:
                    validitytype_list.append('Not Found')
                    
            VALIDITYTYPE_code = pd.Series(validitytype_list, index=self.df.index, name='VALIDITYTYPE_Code')
            return VALIDITYTYPE_code

        for idx, val  in self.df[col_validitytype].items():
            
            if val == '' or pd.isna(val):
                self.audit.warning(f" ValidityType Abscent, Ordre : {idx}, colonne : {col_validitytype}")
                validitytype_list.append('Not Found')
            find = None
            for key, variante  in self.mapping_validitytype.items():
                variante = [x.upper() for x in variante]
                val = val.upper()
                val = val.strip()
                if val in variante :
                    find = key
                    break
            if find :
                self.audit.info(f" ValidityType, Ordre : {idx}, col : {col_validitytype}, code : {val}, Trouvé ")
                validitytype_list.append(find)
            else:
                self.audit.warning(f" ValidityType, Ordre {idx} dans la colonne {col_validitytype} : {val} , Invalide ")
                validitytype_list.append('Not Valid')

        VALIDITYTYPE_code = pd.Series(validitytype_list, index = self.df.index, name ='VALIDITYTYPE_Code')
        return VALIDITYTYPE_code


    def verif_account(self):
        """Verification du comptes indiqués parmi les ceux du client """
        for rules_col in self.dic_rules:
            if self.dic_rules[rules_col]['checkbox']:
                if rules_col == 'Case_unique_account' and self.dic_rules[rules_col]['edit'] != None:
                    return False 
                
        account_list = []

        if 'Account' not in self.correspondances:
            for idx in self.df.index:
                account_list.append('')  
            ACCOUNT_code = pd.Series(account_list, index = self.df.index, name = 'ACCOUNT_Code') 
            return ACCOUNT_code
        
        col_account = self.correspondances ['Account']
        col_account = col_account[0]
        

        # Ajout pour colonne manuelle
        if col_account == "Manuel":
            if 'Account' in list(self.filling_auto.keys()):
                self.audit.info(f" Account, col : {col_account}, code : {self.filling_auto['Account']} ")
                for idx in self.df.index:
                    account_list.append(self.filling_auto['Account'])
                    
            else:
                self.audit.warning(f" Account Abscent (manuel)")
                for idx in self.df.index:
                    account_list.append('Not Found')
                    
            ACCOUNT_code = pd.Series(account_list, index = self.df.index, name = 'ACCOUNT_Code')
            return ACCOUNT_code

    

        Account_ids = list(self.df[col_account].items())
        matched = {}

        for _,elt in Account_ids:
            elt =  self.normalize(elt)
            

            elt = str(elt).upper().strip()
            
            m = []
            
            for acc in self.accounts_dic:
                id = ''
                resultat = process.extract(elt, acc.keys(), scorer = fuzz.partial_ratio)
                
                resultat_cent = [acc[cle] for cle, score, _ in resultat if score == 100]
                if resultat_cent : 
                    id = str(resultat_cent[0])
                    m.append(id)
            matched[elt] = m 
            
        account_dict = {}
        

        for idx, elt in enumerate(matched):
            accounts = matched[elt]
            
            if isinstance(accounts, str):
                
                self.audit.info(f"Ordre : {idx}, Account : {elt} correspondant à {accounts}")
                account_dict[elt] = accounts
                
            elif isinstance(accounts, list):
                unique_accounts = list(set(accounts))
                
                
                if len(unique_accounts) > 1:
                    
                    self.audit.warning(f"Ordre : {idx}, Plusieurs Account différents correspondant à {elt} : {unique_accounts}")
                elif len(unique_accounts) == 1:
                    self.audit.info(f"Ordre : {idx}, Account : {elt} correspondant à {unique_accounts[0]}")
                    account_dict[elt] = unique_accounts[0] 
            else:
                self.audit.warning(f"Ordre : {idx}, Type inattendu pour {elt} : {type(accounts)}")
            
            
                
            

        for idx, val in self.df[col_account].items():
            try:
                val = int(val)
            except:
                val = val
           
            val = self.normalize(val)
            
            
            if pd.isna(val) or val =='':
                self.audit.warning(f" Account Abscent, Ordre : {idx}, colonne : {col_account}")
                account_list.append('Not Found')

            elif val in account_dict :
                
                ul_id = account_dict[val]
                self.audit.info(f" Account, Ordre : {idx}, col : {col_account}, code : {val}, Trouvé, {ul_id} ")
                account_list.append(ul_id)
            else:
                
                self.audit.warning(f" Account, Ordre {idx} dans la colonne {col_account} : {val} , Invalide ou Introuvable ")
                account_list.append('Not Valid')

        ACCOUNT_code = pd.Series(account_list, index = self.df.index, name = 'ACCOUNT_Code')
        
        return ACCOUNT_code


    def verif_price(self):
        """Verification du price si indiqué"""
        price_list = []
  
        if 'OrderType' not in self.correspondances:
            self.audit.info(f" Price Inutile (OrderType absent)")
            for idx in self.df.index:
               
                price_list.append('')
            PRICE_code = pd.Series(price_list, index=self.df.index, name='PRICE_Code', dtype=object)
            return PRICE_code


        col_ordertype = self.correspondances['OrderType']
        col_ordertype = col_ordertype[0]


        
        if col_ordertype == 'Manuel':
            if 'Price' not in self.correspondances:
                self.audit.warning(f" Price Abscent (manuel)")
                for idx in self.df.index:
                    price_list.append('')
                    
                PRICE_code = pd.Series(price_list, index = self.df.index, name = 'PRICE_Code',dtype = object)
                return PRICE_code
            else:
                col_price = self.correspondances ['Price']
                col_price = col_price[0]
                if col_price == 'Manuel':
                    if 'Price' in list(self.filling_auto.keys()):
                        self.audit.info(f" Account, col : {col_price}, code : {self.filling_auto['Price']} ")
                        for idx in self.df.index:
                            price_list.append(self.filling_auto['Price'])
                           
                    else:
                        self.audit.warning(f" Price Abscent (manuel), Ordre : {idx}")
                        for idx in self.df.index:
                            price_list.append('Not Found')
                    PRICE_code = pd.Series(price_list, index = self.df.index, name = 'PRICE_Code')
                    return PRICE_code
                else:
                    for idx, val in self.df[col_price].items():
                        if val == '' or pd.isna(val):
                            if self.ORDERTYPE.iloc[idx].strip() == 'limit' or self.ORDERTYPE.iloc[idx].strip() == 'stoplimit':
                                self.audit.warning(f" Price Abscent, Ordre : {idx}, colonne : {col_price}")
                                price_list.append('Not Found')
                            else:
                                self.audit.info(f" Price Inutile, Odre {idx}, col : {col_price}")
                                price_list.append('')
                        else:
                            val = self.to_float(val)
                            
                            if val != None  :
                                self.audit.info(f" Price, Ordre : {idx}, col : {col_price}, code : {val}, Trouvé ")
                                price_list.append(val)
                            else:
                                self.audit.warning(f" Price, Ordre {idx} dans la colonne {col_price} : {val} , Invalide ") 
                                price_list.append('Not Valid')
                    PRICE_code = pd.Series(price_list, index = self.df.index, name = 'PRICE_Code',dtype = object)
                        
                    return PRICE_code



        
        if 'Price' not in self.correspondances:
            for idx, val in self.df[col_ordertype].items():
                if self.ORDERTYPE.iloc[idx].strip() == 'limit' or self.ORDERTYPE.iloc[idx].strip() == 'stoplimit':
                    self.audit.warning(f" Price Abscent, Ordre : {idx}, colonne : {col_price}")
                    price_list.append('Not Found')
                else:
                    self.audit.info(f" Price Inutile, Odre {idx} ")
                    price_list.append('')

        else:     
            col_price = self.correspondances ['Price']
            col_price = col_price[0]

            if col_price not in self.df.columns:
                self.audit.error(f"La colonne {col_price} n'existe pas dans le BasketTrade")
                raise ValueError(f"La colonne {col_price} n'existe pas dans le BasketTrade" )
            

            for idx, val in self.df[col_price].items():
                if val == '' or pd.isna(val):
                    if self.ORDERTYPE.iloc[idx].strip() == 'limit' or self.ORDERTYPE.iloc[idx].strip() == 'stoplimit':
                        self.audit.warning(f" Price Abscent, Ordre : {idx}, colonne : {col_price}")
                        price_list.append('Not Found')
                    else:
                        self.audit.info(f" Price Inutile, Odre {idx}, col : {col_price}")
                        price_list.append('')
                else:
                    val = self.to_float(val)
                    
                    if val != None  :
                        self.audit.info(f" Price, Ordre : {idx}, col : {col_price}, code : {val}, Trouvé ")
                        price_list.append(val)
                    else:
                        self.audit.warning(f" Price, Ordre {idx} dans la colonne {col_price} : {val} , Invalide ") 
                        price_list.append('Not Valid')

        PRICE_code = pd.Series(price_list, index = self.df.index, name = 'PRICE_Code',dtype = object)
        
        return PRICE_code

    def verif_stopprice(self):
        """ Verfication du stopprice si indiqué """
        stopprice_list = []

        if 'OrderType' not in self.correspondances:
            self.audit.info(f" StopPrice Inutile (OrderType absent)")
            for idx in self.df.index:
                
                stopprice_list.append('')
            STOPPRICE_code = pd.Series(stopprice_list, index=self.df.index, name='STOPPRICE_Code', dtype=object)
            return STOPPRICE_code
        
        col_ordertype = self.correspondances['OrderType']
        col_ordertype = col_ordertype[0]


        if col_ordertype == 'Manuel':
            if 'StopPrice' not in self.correspondances:
                self.audit.warning(f" StopPrice Abscent (manuel)")
                for idx in self.df.index:
                    stopprice_list.append('')
                   
                STOPPRICE_code = pd.Series(stopprice_list, index = self.df.index, name = 'STOPPRICE_Code',dtype = object)
                return STOPPRICE_code
            else:
                col_stopprice = self.correspondances ['StopPrice']
                col_stopprice = col_stopprice[0]
                
                if col_stopprice == 'Manuel':
                    
                    if 'StopPrice' in list(self.filling_auto.keys()):
                        self.audit.info(f" StopPrice, col : {col_stopprice}, code : {self.filling_auto['StopPrice']} ")
                        for idx in self.df.index:
                            
                            stopprice_list.append(self.filling_auto['StopPrice'])
                            

                    else:
                        for idx, val in self.df[col_stopprice].items():    
                            if val == '' or pd.isna(val):
                                if self.ORDERTYPE.iloc[idx] == 'stop' or self.ORDERTYPE.iloc[idx] == 'stoplimit':
                                    self.audit.warning(f" StopPrice Abscent, Ordre : {idx}, colonne : {col_stopprice}")
                                    stopprice_list.append('Not Found')
                                else: 
                                    self.audit.info(f" StopPrice Inutile, Odre {idx}, col : {col_stopprice}")
                                    stopprice_list.append('')
                            else: 
                                    val = self.to_float(val)
                                    if val != None  :
                                        self.audit.info(f" StopPrice, Ordre : {idx}, col : {col_stopprice}, code : {val}, Trouvé ")
                                        stopprice_list.append(val)
                                    else:
                                        self.audit.warning(f" StopPrice, Ordre {idx} dans la colonne {col_stopprice} : {val} , Invalide ") 
                                        stopprice_list.append('Not Valid')
                    STOPPRICE_code = pd.Series(stopprice_list, index = self.df.index, name = 'STOPPRICE_Code',dtype = object)
                    return STOPPRICE_code

        

        if 'StopPrice' not in self.correspondances:
            
            for idx, val in self.df[col_ordertype].items():
                if self.ORDERTYPE.iloc[idx] == 'stop' or self.ORDERTYPE.iloc[idx] == 'stoplimit':
                    self.audit.warning(f" StopPrice Abscent, Ordre : {idx}, colonne : {col_stopprice}")
                    stopprice_list.append('Not Found')
                else:
                    self.audit.info(f" StopPrice Inutile, Odre {idx} ")
                    stopprice_list.append('')

        else:        
            col_stopprice = self.correspondances ['StopPrice']
            col_stopprice = col_stopprice[0]

            if col_stopprice not in self.df.columns:
                self.audit.error(f"La colonne {col_stopprice} n'existe pas dans le BasketTrade")
                raise ValueError(f"La colonne {col_stopprice} n'existe pas dans le BasketTrade" )
            
            for idx, val in self.df[col_stopprice].items():
                
                if val == '' or pd.isna(val):
                    if self.ORDERTYPE.iloc[idx] == 'stop' or self.ORDERTYPE.iloc[idx] == 'stoplimit':
                        self.audit.warning(f" StopPrice Abscent, Ordre : {idx}, colonne : {col_stopprice}")
                        stopprice_list.append('Not Found')
                    else: 
                        self.audit.info(f" StopPrice Inutile, Odre {idx}, col : {col_stopprice}")
                        stopprice_list.append('')
                else: 
                        val = self.to_float(val)
                        if val != None  :
                            self.audit.info(f" StopPrice, Ordre : {idx}, col : {col_stopprice}, code : {val}, Trouvé ")
                            stopprice_list.append(val)
                        else:
                            self.audit.warning(f" StopPrice, Ordre {idx} dans la colonne {col_stopprice} : {val} , Invalide ") 
                            stopprice_list.append('Not Valid')
                    



        STOPPRICE_code = pd.Series(stopprice_list, index = self.df.index, name = 'STOPPRICE_Code',dtype = object)
        
        return STOPPRICE_code

    def verif_expiredate(self):
        """ Verification de la date d'expiration si indiqué """
        expiredate_list = []

        if 'ValidityType' not in self.correspondances:
            self.audit.info(f" ExpireDate Inutile (ValidityType absent)")
            for idx in self.df.index:
                
                expiredate_list.append('')
            EXPIREDATE_code = pd.Series(expiredate_list, index=self.df.index, name='EXPIREDATE_Code', dtype=object)
            return EXPIREDATE_code

        col_validitytype = self.correspondances['ValidityType']
        col_validitytype = col_validitytype[0]

        if col_validitytype == 'Manuel':
            self.audit.warning(f" ExpireDate Abscent (manuel)")
            for idx in self.df.index:
                expiredate_list.append('')
                
            EXPIREDATE_code = pd.Series(expiredate_list, index = self.df.index, name = 'EXPIREDATE_Code',dtype = object)
            return EXPIREDATE_code

        if 'ExpireDate' not in self.correspondances:
            for idx, val in self.df[col_validitytype].items():
                if self.VALIDITYTYPE.iloc[idx] == 'gtd':
                    self.audit.warning(f" ExpireDate Abscent, Ordre : {idx}")
                    expiredate_list.append('Not Found')
                else:
                    self.audit.info(f" ExpireDate Inutile, Odre {idx} ")
                    expiredate_list.append('')
        else:        
            col_expiredate = self.correspondances ['ExpireDate']
            col_expiredate = col_expiredate[0]

            if col_expiredate not in self.df.columns:
                self.audit.error(f"La colonne {col_expiredate} n'existe pas dans le BasketTrade")
                raise ValueError(f"La colonne {col_expiredate} n'existe pas dans le BasketTrade" )
            
            for idx, val in self.df[col_expiredate].items():
                if val == '' or pd.isna(val):
                    if self.VALIDITYTYPE.iloc[idx] == 'gtd':
                        self.audit.warning(f" ExpireDate Abscent, Ordre : {idx}, colonne : {col_expiredate}")
                        expiredate_list.append('Not Found')
                    else:
                        self.audit.info(f" ExpireDate Inutile, Odre {idx}, col : {col_expiredate}")
                        expiredate_list.append('')
                else:

                    val = self.to_yyyymmdd(val)
                    if val != None  :
                            self.audit.info(f" ExpireDate Inutile, Odre {idx}, col : {col_expiredate}")
                            expiredate_list.append(val)
                    else:
                            self.audit.info(f" Price, Ordre : {idx}, col : {col_expiredate}, code : {val}, Trouvé ")
                            expiredate_list.append('Not Valid')



        EXPIREDATE_code = pd.Series(expiredate_list, index = self.df.index, name = 'EXPIREDATE_Code')
        
        return EXPIREDATE_code
      

    def verif_reference(self):
        """Allocation de la référence fourni ou création d'une nouvelle"""
        reference_list = []


        if 'Reference' in self.correspondances:
            col_reference = self.correspondances ['Reference']
            col_reference = col_reference[0]
            for idx, val in self.df[col_reference].items() :
                if val !='' and pd.notna(val):
                    self.audit.info(f"Reference Presente, Ordre : {idx}, col : {col_reference}, {val}") 
                    reference_list.append(val)
                else:
                    date = datetime.today().strftime("%Y%m%d")
                    val = f"{date}-{idx}"
                    self.audit.info(f"Reference Ajout, Ordre : {idx},  col : {col_reference}, {val}")
                    reference_list.append(val)                
        else: 
            for idx in range(len(self.df)):
                    date = datetime.today().strftime("%Y%m%d")
                    val = f"{date}-{idx}"
                    self.audit.info(f"Reference Ajout, Ordre : {idx}, Nouvelle colonne , {val}") 
                    reference_list.append(val)


        REFERENCE_code = pd.Series(reference_list, index = self.df.index, name ='REFERENCE_Code')

        return REFERENCE_code

    def verif_comment(self):
        """Allocation d'un commentaire pouvant être une stratégie, une indication ... """
        comment_list = []

        
        if 'Comment' in self.correspondances:
            col_comment = self.correspondances ['Comment']

            if len(col_comment) == 1 :   
                col_comment = col_comment[0]
                if col_comment == 'Manuel':
                    if 'Comment' in list(self.filling_auto.keys()):
                        comment_list.append(self.filling_auto['Comment'])
                        self.audit.info(f"Comment Manuel , col : {col_comment} ")

                else:
                    for idx, val in self.df[col_comment].items() :
                        if val != '' and pd.notna(val):
                            self.audit.info(f"Comment Present, Order : {idx} , col : {col_comment}, {val} ") 
                            comment_list.append(val)
                        else :
                            comment_list.append('')
                            self.audit.info(f"Comment Abscent, Order : {idx} , col : {col_comment} ")

            elif len(col_comment) >1:
                self.verif_darkpool(col_comment)
                self.df['Comment_fusion'] = (self.df[col_comment].apply(lambda row: ' / '.join([str(x) for x in row if pd.notna(x) and x !='']), axis =1))
                for idx, val in self.df['Comment_fusion'].items() :
                        if val != '' and pd.notna(val):
                            self.audit.info(f"Comment Present, Order : {idx} , col : {col_comment}, {val} ") 
                            comment_list.append(val)
                        else :
                            comment_list.append('')
                            self.audit.info(f"Comment Abscent, Order : {idx} , col : {col_comment} ")

        else:
            for idx in range(len(self.df)):
                self.audit.info(f"Comment Abscent , Ordre : {idx}, Nouvelle colonne ") 
                comment_list.append('')


        COMMENT_code = pd.Series(comment_list, index = self.df.index, name ='COMMENT_Code')

        return COMMENT_code


    #-----------------------------------------

    def verif_darkpool(self,col_comment):
        """Fonction prenant en entrée la colonne commentaire et indentifiant la présence d'une indication de darkpool  """
        target = 'dark pool'
        best_match = process.extractOne(target,col_comment, scorer = fuzz.token_sort_ratio)

        if best_match[1] >= 50:
           best_match_col = process.extractOne(target,self.df.columns.tolist(), scorer = fuzz.token_sort_ratio)
           self.df[best_match_col[0]] = ' Dark pool : ' + self.df[best_match_col[0]].fillna('NaN').astype(str)
           col_comment[best_match[2]] = best_match_col[0]
        return False 


    #-----------------------------------------

    def extract_isin(self,val):
        """ Fromatage de l'isin fourni sous la forme FR0000000000"""
        match = re.search(r'\b[A-Z]{2}[A-Z0-9]{9}[0-9]\b',val)
        return match.group(0) if match else None 

    def extract_bloom(self,val):
        """Mise en forme du ticker bloom dans un format sans mension "Equity" """
        val = str(val)
        val = val.strip().upper()
        if f"\xa0" in val:
            val = val.replace(f"\xa0", ' ')
        if "EQUITY" in val:
            val =  val.replace('EQUITY','')
        l = ['-',':',';','.']
        for i in l:
            if i in val:
                val = val.replace(i,' ')
        
        return val.strip()

    def msg_isin_bloom(self):
        """Affiche un message WARNING en cas de ISIN non correspondant"""
        if self.messagecount:
            # Créer une fenêtre d'avertissement non-modale
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("Codes Isin incorrect")
            msg.setText(self.message)
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.setModal(False)  # Non-modal = ne bloque pas
            msg.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
            msg.show()  # show() au lieu de exec()

            # Garder une référence pour éviter la suppression automatique
            self.Isin_warning_msg = msg
 

    def is_float(self,s):
        """Convertisseur d'un str en float si possible"""
        try:
            float(s)
            return True
        except ValueError:
            return False

        

    def to_float(self,val):
        if isinstance(val,(int,float)):
            return float(val)
        elif isinstance(val,str):
            try:
                val = float(val)
                return val
            except ValueError:
                return None
        return None

    def to_int(self,val):
        """Convertisseur de float et str en entier quitte à utiliser la fonction "round" """
    
        if isinstance(val,int):
            val = abs(val)
            
            val = int(val)
            
            return val
        elif isinstance(val,float):
            val = abs(val)
            
            x = round(val)
            
            if x >= 0:
                return x 
            else:
                return None
        elif isinstance(val, str):
            try:
                
                val = val.replace(" ","")
                match = re.search(r'\d+',val)

                return match.group(0) if match else None
            except  ValueError:
                return None
        return None

    def to_yyyymmdd(self,val):
        """Formatage de la date en format YYYYMMDD"""
        if isinstance(val,(int,float)):
            val = str(int(val))
        elif not isinstance(val,str):
            return None 
        
        val = val.strip()

        formats = [
            "%Y-%m-%d",
            "%Y-%m-%d %H:%M:%S",
            "%Y%m%d",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d.%m.%Y",
            "%d/%m/%y",
            "%d-%m-%y",
            "%d.%m.%y",
        ]
        for fmt in formats:
            try:
                dt = datetime.strptime(val,fmt)
                self.audit.info(f"Formatage date de {val} à {dt}")
                if "%y" in fmt : 
                    year = dt.year
                    if year < 2000:
                        dt = dt.replace(year = year + 2000)
                return dt.strftime("%Y%m%d")
            except ValueError:
                continue 
        return None 

    def normalize(self,x):
        """Normalisation en str"""
        if x is None:
            return None
        elif x == 'NaN':
            return x 
        
        if isinstance(x,(float,int)):
            x_new = str(int(x))
        elif isinstance(x,str):
            x_new = x
            x_clean = x.replace(",",".")
            if re.fullmatch(r"\d+(\.\d+)?", x_clean):
                x_new = str(int(float(x_clean)))
        
        if 'CACEIS' in x_new:
            x_new = x_new.replace('CACEIS','').strip()
        return x_new
            
    #-----------------------------------------
    
    def complete_bloom_columns(self):
        
        """Si une colonne Bloomberg est présente :
                - Cherche le produit dans bloomberg
                - Compare l'ISIN du produit avec celui donnée
                - Affiche en cas d'incohérance
                - Crée les colonnes Exchange et Currency si elles ne sont pas données"""
        
        if self.no_isin:
            if 'Currency' not in self.correspondances:
                self.df['Currency'] = None 
                self.correspondances['Currency'] = ['Currency']
                

                for code, info in self.only_bloom_info.items():
                   

                    self.df.iloc[info["idx"],self.df.columns.get_loc("Currency")] = info['currency']     
                             
                    self.audit.info(f"Ordre : {info['idx']} ajout Currency {info['currency']}")
                

           
            self.df['Exchange'] = None 
            self.correspondances['Exchange'] = ['Exchange']

            for code, info in self.only_bloom_info.items():
                exchange = info['exchange']
                if  exchange in self.mapping_bloom:
                    exchange = self.mapping_bloom[exchange]
                self.df.iloc[info["idx"],self.df.columns.get_loc("Exchange")] = exchange
                self.audit.info(f"Ordre : {info['idx']} ajout Exchange {info['exchange']}")
            
            
            if 'Currency' in self.correspondances:
                col_cur = self.correspondances['Currency'][0]
                

                self.df['Currency_int'] = None
                 
                for code, info in self.only_bloom_info.items():
                    self.df.iloc[info["idx"],self.df.columns.get_loc("Currency_int")] = info['currency']
                
                self.df["Currency_comparison"] = np.where(self.df[col_cur].str.upper().str.strip()== self.df["Currency_int"].str.upper().str.strip(),self.df[col_cur].str.upper().str.strip(), 'Not Valid')
                self.audit.info(f" Currency après la compraison {self.df["Currency_comparison"]} ")
                self.df[col_cur] = self.df["Currency_comparison"] 




        else:
            # Cherche la colonne Bloomberg dans le mapping
            if 'Bloomberg Code' not in self.correspondances:
                self.audit.info('Colonne Bloomberg Abscent')
                return False
            else :

                for key in ['Currency','Exchange']:
                            if key not in self.correspondances:
                                    self.correspondances[key] = [key]
                                    for idx in range(len(self.df)):
                                        self.df.at[idx,key] = ''

                # Si colonne trouvée, mapping comme avant
                col_bloom = self.correspondances['Bloomberg Code'][0]
        

                bloom_liste = self.df[col_bloom].tolist()

                for idx in range(len(bloom_liste)):
                    bloom_liste[idx] = self.extract_bloom(bloom_liste[idx])
                    
                    bloom_liste[idx] = bloom_format_transformation_ticker(bloom_liste[idx])
                    self.compare_isin(bloom_liste[idx],idx)
                    
                        
                    if idx in self.isin_tickerbloomformat_dic:
                        for key in ['Currency','Exchange']:
                                isin = self.isin_tickerbloomformat_dic[idx]['isin']
                                ticker = self.isin_tickerbloomformat_dic[idx]['ticker']
                                        
                                if key == 'Exchange':
                                    self.df.at[idx,key] = self.isin_info[isin][ticker]['exchange']
                                    self.audit.info(f"Ordre {idx} prend la valeur {self.isin_info[isin][ticker]['exchange']} dans la colonne {key} ")
                                else: 
                                    self.df.at[idx,key] = self.isin_info[isin][ticker]['currency']
                                    self.audit.info(f"Ordre {idx} prend la valeur {self.isin_info[isin][ticker]['currency']} dans la colonne {key} ")
                    else:
                        self.df.at[idx,key] = ''
                        
                        

        
        return True


    def compare_isin(self,ticker_bloom,idx):
        """Compare l'Isin du Bloomberg et celui du produit donnée """
        l = []
        colonne_isin = self.correspondances['isin']
        colonne_isin = colonne_isin[0]
        isin_fourni =  self.df.iloc[idx][colonne_isin]
        isin_fourni = self.extract_isin(isin_fourni)
        if self.isin_info[isin_fourni] :
            if ticker_bloom not in self.isin_info[isin_fourni]:
                if idx not in self.idx_isin:

                    self.messagecount = True
                    
                    self.idx_isin[idx] = [isin_fourni,ticker_bloom]
                    self.isin_comparison[isin_fourni] = [ticker_bloom,idx]
                    
                    self.ticker_unknown_isin_list.append(ticker_bloom)
                    self.idx_get_ticker_info_from_tickers.append(idx)
                    # Construire le message
                    
                    self.message += f"• ISIN Fourni : {isin_fourni} / Bloomberg : {ticker_bloom}\n"

                    # Logs audit
                    self.audit.warning("Codes Isin incorrect")
                    
                    self.audit.warning(f" ISIN Fourni : {isin_fourni} / Bloomberg : {ticker_bloom}")
            else : 
                self.isin_tickerbloomformat_dic[idx] = {'ticker' : ticker_bloom, 'isin' : isin_fourni}
                
        
        else:
            
            self.ticker_unknown_isin_list.append(ticker_bloom)
            self.idx_get_ticker_info_from_tickers.append(idx)
            self.messagecount = True
            self.idx_isin[idx] = [isin_fourni,ticker_bloom]
            self.audit.warning(f" ISIN Fourni : {isin_fourni} Introuvable")
            self.message += f"• ISIN Fourni : {isin_fourni} Introuvable \n"
           



      
    def affichage_notfoundBloom(self):
        """Affiche les erreurs causées par le ticker bloomberg"""
        
        if not self.notfoundBloom:
            return

        # Construire le message
        message = f"{len(self.notfoundBloom)} code(s) Bloomberg non trouvé(s) :\n\n"
        for bloom_code, details in self.notfoundBloom.items():
            message += f"• {bloom_code} : {details}\n"

        message += "\nVeuillez vérifier ces codes Bloomberg."

        # Créer une fenêtre d'avertissement non-modale
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("Codes Bloomberg non trouvés")
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.setModal(False)  # Non-modal = ne bloque pas
        msg.show()  # show() au lieu de exec()

        # Garder une référence pour éviter la suppression automatique
        self.bloomberg_warning_msg = msg

        # Logs audit
        self.audit.info("Codes Bloomberg non trouvés")
        for bloom_code, details in self.notfoundBloom.items():
            self.audit.info(f" - {bloom_code} : {details}")

    #-----------------------------------------

    def col_client_memory(self):
        """ Crée la colonne col_client_memory_list qui est composé des colonnes de la table TBL_BTRADE_CLIENTMEMORY """

        conn = pymssql.connect(
        server = '',
        user = "",
        password = '',
        database = ''
            ) 
        cursor = conn.cursor()
        query_col= "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TBL_BTRADE_CLIENTMEMORY' ORDER BY ORDINAL_POSITION "
        cursor.execute(query_col)
        rows = cursor.fetchall()
        self.col_client_memory_list = [row[0]for row in rows if row[0] is not None ]
        conn.close()

        self.audit.info(f"Information en memoire du client : {self.col_client_memory_list}")

        self.info_rules[self.col_client_memory_list[18]] = None
        
        for col in self.col_client_memory_list[20:]:
            self.info_rules[col] = None

    #-----------------------------------------
 
    def rules(self):
        """Applique les règles données dans verifWindow"""

        for rules_col in self.dic_rules:
            if self.dic_rules[rules_col]['checkbox']:
                if rules_col == 'Case_unique_account' and self.dic_rules[rules_col]['edit'] != None:
                    self.ACCOUNT = self.Case_unique_account(self.df, self.dic_rules[rules_col]['edit'],self.accounts_dic)
                elif rules_col == 'unique_currency' and self.dic_rules[rules_col]['edit'] != None:
                    self.CURRENCY = self.unique_currency(self.df, self.dic_rules[rules_col]['edit'])
                elif rules_col == 'unique_exchange' and self.dic_rules[rules_col]['edit'] != None:
                    self.EXCHANGE = self.unique_exchange(self.df, self.dic_rules[rules_col]['edit'])
                elif rules_col == 'supress_order_on_quantity' and self.dic_rules[rules_col]['edit'] != None:
                    self.supress_order_auto()
                elif rules_col == 'maximum_volume_exchange' and self.dic_rules[rules_col]['edit'] != None:
                    self.maximum_volume_exchange()

        
    def Case_unique_account(self,df, data,accounts_dic):
        """Règle : compte du client dans une seule case de l'excel"""

        val = self.excel_extract_value(data)
        
        val = str(val).upper().strip()
        m = []
        matched = {}  
        account_list = []    
        for acc in accounts_dic:
            id = ''
            resultat = process.extract(val, acc.keys(), scorer = fuzz.partial_ratio)
            resultat_cent = [acc[cle] for cle, score, _ in resultat if score == 100]
            if resultat_cent : 
                id = str(resultat_cent[0])
                m.append(id)
        matched[val] = m 
                    
        account_dict = {}
            

        for idx, elt in enumerate(matched):
            accounts = matched[elt]
                
            if isinstance(accounts, str):
                account_dict[elt] = accounts
                    
            elif isinstance(accounts, list):
                unique_accounts = list(set(accounts))
                    
                if len(unique_accounts) == 1:
                        account_dict[elt] = unique_accounts[0] 

                    
        for l in range(len(df)):
            try:
                val = int(val)
            except:
                val = val
                val = str(val).strip()
                
                
            if pd.isna(val) or val =='':
                self.audit.warning(f"Case_unique_account, position: {data}, Account : Not Found")
                account_list.append('Not Found')

            elif val in account_dict :
                self.audit.info(f"Case_unique_account, position: {data}, ,Account : {val}")
                ul_id = account_dict[val]
                account_list.append(ul_id)
            else:
                account_list.append('Not Valid')
                self.audit.info(f"Case_unique_account, position: {data}, Not Valid")

        ACCOUNT_code = pd.Series(account_list, index = df.index, name = 'ACCOUNT_Code')
        return ACCOUNT_code

    def unique_currency(self,df,data):
        """Regle : le client traite avec une seule currency"""
        currency_list = []
        self.audit.info(f"Unique Currency : {data}")
        for l in range(len(df)):
            currency_list.append(data)

        CURRENCY_code = pd.Series(currency_list, index = df.index, name = 'CURRENCY_Code')
        return CURRENCY_code

    def unique_exchange(self,df,data):
        """Regle: le client traite sur une seule place"""

        self.audit.info(f"Unique Exchange : {data}")
        exchange_list = []
        for l in range(len(df)):
            exchange_list.append(data)
        EXCHANGE_code = pd.Series(exchange_list, index = df.index, name = 'EXCHANGE_Code')
        return EXCHANGE_code



    
        #-----------------------------------------

    def supress_order_auto(self):
        """Supprime automatiquement les ordres ne donnant aucune quantité valide"""
        l = ['Not Found', 'Not Valid', '0',0,' ', '']
        supress_order_idx = self.QUANTITY[self.QUANTITY.isin(l)].index.tolist()
        self.QUANTITY= self.QUANTITY[~self.QUANTITY.isin(l)].reset_index(drop = True)

        self.ISIN = self.ISIN.drop(supress_order_idx).reset_index(drop=True)
        self.EXCHANGE = self.EXCHANGE.drop(supress_order_idx).reset_index(drop=True)
        self.SIDE = self.SIDE.drop(supress_order_idx).reset_index(drop=True)
        self.CURRENCY = self.CURRENCY.drop(supress_order_idx).reset_index(drop=True)
        self.ORDERTYPE = self.ORDERTYPE.drop(supress_order_idx).reset_index(drop=True)
        self.VALIDITYTYPE = self.VALIDITYTYPE.drop(supress_order_idx).reset_index(drop=True)
        self.ACCOUNT = self.ACCOUNT.drop(supress_order_idx).reset_index(drop=True)
        self.PRICE = self.PRICE.drop(supress_order_idx).reset_index(drop=True)
        self.STOPPRICE = self.STOPPRICE.drop(supress_order_idx).reset_index(drop=True)
        self.EXPIREDATE = self.EXPIREDATE.drop(supress_order_idx).reset_index(drop=True)
        self.REFERENCE = self.REFERENCE.drop(supress_order_idx).reset_index(drop=True)
        self.COMMENT = self.COMMENT.drop(supress_order_idx).reset_index(drop=True)

    def maximum_volume_exchange(self):
        """Remplit automatiquement les exchanges à l'aide du ticker bloom avec le plus grand volume"""
        
        for idx in range(len(self.ISIN)):
                isin = self.ISIN.iloc[idx]
                currency_fourni = self.CURRENCY[idx]
                if currency_fourni not in ('Not Found', 'Not Valid', ''):
                    dic_exchange_brut = self.isin_info[isin]
                    if dic_exchange_brut:
                        liste_couple_exchange_currency = []
                        liste_volume = []
                        for d in dic_exchange_brut.keys():
                            exchange = dic_exchange_brut[d]['exchange']
                            currency_bloom = dic_exchange_brut[d]['currency']
                            if not exchange:
                                exchange = 'Not Valid'
                            if not currency_bloom:
                                currency_bloom = 'Not Valid'
                            if currency_fourni.upper() == currency_bloom.upper():
                                liste_couple_exchange_currency += [[exchange,currency_fourni]]
                                liste_volume.append(int(dic_exchange_brut[d]['volume_moyen'].replace(',','')))
                        if liste_volume != []:
                            index, value = max(enumerate(liste_volume),key  = lambda x : x[1])
                            exchange = liste_couple_exchange_currency[index][0].strip().upper()
                            if exchange in self.mapping_bloom:
                                exchange = self.mapping_bloom[exchange]

                            self.EXCHANGE[idx] = exchange
                            
                            self.audit.info(f"MODIFICATION EXCHANGE/CURRENCY MAXIMUM FONCTION MEMORY: row = {idx}, Exchange = {exchange} / Currency = {liste_couple_exchange_currency[index][1].strip().upper()} ")


# =================================================================================================
# =                                           SECTION: VERIFICATION 
# =================================================================================================


class CorrectionWindow (QWidget):
    """Ouvre une fenetre de corresction avant envoie"""

    def __init__(self, FILE, fclose):
        super().__init__()
        self.force = False 
        self.fclose = fclose
        self.FILE = FILE
        self.audit = self.FILE.audit
        self.idx_isin = FILE.idx_isin
        self.client_name = FILE.clientUlID
        self.info_bloom_dic = FILE.isin_info
        self.ticker_unknown_isin = FILE.ticker_unknown_isin
        self.client_accounts = FILE.client_accounts
        self.ordered_attrs = ["ISIN","EXCHANGE","CURRENCY","SIDE","QUANTITY","ORDERTYPE","PRICE",'STOPPRICE','VALIDITYTYPE','EXPIREDATE','ACCOUNT','REFERENCE','COMMENT']
        self.df = self._build_dataframe()
        self.FILE.audit.info("DataFrame entrée : \n%s", self.df.to_string(index=False))
        self.exchange_incorrecte = self.FILE.exchange_incorrecte
        self.mapping_bloom = FILE.mapping_bloom
        self.etf = FILE.etf
        self.no_isin = FILE.no_isin
        self.message_no_isin = FILE.message_no_isin
        self.use_info = FILE.Use_info

        self.setWindowTitle("Correction")
        self.setGeometry(100, 100, 1000, 600)
        self.setFixedSize(1100,600)

        layout = QVBoxLayout()
        
        self.table = QTableWidget()
        self.table.setRowCount(len(self.df))
        self.table.setColumnCount(len(self.df.columns))
        self.table.setHorizontalHeaderLabels(self.df.columns)

        # Configuration du header pour le menu contextuel
        horizontal_header = self.table.horizontalHeader()
        horizontal_header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        horizontal_header.customContextMenuRequested.connect(self.show_column_context_menu)

        self.table.verticalHeader().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.verticalHeader().customContextMenuRequested.connect(self.delete_order)

        # Affiche le message d'erreur qui apparait quand il n'y a pas de colonne ISIN est qu'au moins un ticker introuvable 
        if self.message_no_isin != '':
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("Ticker introuvable")
            msg.setText(self.message_no_isin)
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.setModal(False)  # Non-modal = ne bloque pas
            msg.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
            msg.show()  # show() au lieu de exec()
            self.msg_no_isin = msg


        for row in range(len(self.df)):
            for col, col_name in enumerate(self.df.columns):
                val = self.df.iloc[row, col]
                text = str(val) if pd.notna(val) else " "
                item = QTableWidgetItem(text)
                
                 
                if self.no_isin == False :
                    if (row in self.idx_isin and col == 0) :
                        if self.info_bloom_dic[item.text()]:
                            item.setBackground(QColor(176,9,27))
                        else:
                            item.setBackground(QColor(232, 78, 15))
                    
                    elif (row in self.exchange_incorrecte and col == 1):
                        item.setBackground(QColor(195,204,213))
                        
                    elif text == '' and col != 10: 
                        item.setBackground(QColor(255, 255, 200))
                    elif text == "Not Found" or text == 'Not Valid' or text == 'ISIN Invalide'  or (col == 10 and text == '')  :
                        item.setBackground(QColor(255, 150, 150))
                    else: 
                        item.setBackground(QColor(240, 240, 240))
                else:
                    if col == 0 and text == 'FR0000000000':
                        item.setBackground(QColor(176,9,27))
                    elif text == '' and col != 10: 
                        item.setBackground(QColor(255, 255, 200))
                    elif text == "Not Found" or text == 'Not Valid' or text == 'NOT VALID' or (col == 10 and text == '')  :
                        item.setBackground(QColor(255, 150, 150))
                    else: 
                        item.setBackground(QColor(240, 240, 240))

                self.table.setItem(row, col, item)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.on_right_click)
        layout.addWidget(self.table)

        bottom_layout = QHBoxLayout()

        self.correctionbutton = CorrectionButton(self.client_name)
        self.correctionbutton.clicked.connect(self.correctionbutton.open_correction)
        bottom_layout.addWidget(self.correctionbutton)

        self.save_button = QPushButton("Valider")
        self.save_button.clicked.connect(self.save_correction)
        bottom_layout.addWidget(self.save_button)

        layout.addLayout(bottom_layout)

        self.setLayout(layout)


    def delete_order(self,pos: QPoint):
        """Supprime les ordres avec une quantité invalide"""
        header = self.table.verticalHeader()
        row = header.logicalIndexAt(pos)

        if row <0 :
            return 
        item = self.table.item(row,4)
        value = item.text() if item else None 
        
        if value not in ['Not Found', '0', '']:
            return
        menu = QMenu(self)
        delete_action = menu.addAction("Supprimer l'ordre")
        action = menu.exec(header.mapToGlobal(pos))

        if action == delete_action:
            self.table.removeRow(row)
            self.audit.info(f"Order {row} : Ordre supprimé")


    # Menu clic droit 
    def on_right_click(self,pos: QPoint):
        """Option de click droit suivant la colonne selectionné"""
        item = self.table.itemAt(pos)
        if item is None:
            return
        
        row = item.row()
        column = item.column()
        value = item.text()
        menu = QMenu(self)
        isin = self.table.item(row,0)
        
        if column == 0 :
            if self.no_isin :
                return False 
           
            if row in self.idx_isin:
                liste_info = []

                for ticker in self.ticker_unknown_isin:
                    if self.ticker_unknown_isin[ticker]:
                        if self.ticker_unknown_isin[ticker]['idx'] == row:
                            if 'isin' in self.ticker_unknown_isin[ticker] and 'volume_moyen' in self.ticker_unknown_isin[ticker] :
                                isin = self.ticker_unknown_isin[ticker]['isin']
                                currency = self.ticker_unknown_isin[ticker]['currency']
                                volume_moyen_30 = self.ticker_unknown_isin[ticker]['volume_moyen']
                                exchange = self.ticker_unknown_isin[ticker]['exchange']
                                name =  self.ticker_unknown_isin[ticker]['name']
                                if not exchange:
                                    exchange = 'Not Valid'
                                if not currency:
                                    currency = 'Not Valid'
                                chaine =  isin + '//' + ticker + ' // ' + exchange + ' // ' + currency +  ' // '  + name + '//' + volume_moyen_30 
                                liste_info.append(chaine)
                
                actions = {menu.addAction(acc): acc for acc in liste_info}
                action = menu.exec(self.table.viewport().mapToGlobal(pos))

                if action in actions:
                    selected_account = actions[action]
                    selected_account = selected_account.split('//')
                    self.table.setItem(row,column, QTableWidgetItem(selected_account[0].strip()))
                    exchange = selected_account[2].strip()
                    for idx in range(self.table.rowCount()):
                        if self.table.item(idx, 0).text() == isin :
                            if exchange in self.mapping_bloom:
                                exchange = self.mapping_bloom[exchange]
                            self.table.setItem(idx,column + 1, QTableWidgetItem(exchange))
                            self.table.setItem(idx,column + 2, QTableWidgetItem(selected_account[3].strip()))
                            self.audit.info(f"MODIFICATION ISIN/EXCHANGE/CURRENCY: row = {idx}, ISIN = {selected_account[3].strip()} / Exchange = {selected_account[1].strip()} / Currency = {selected_account[2].strip()} ")


               
        if column == 3 : 
            
            action_buy = menu.addAction("BUY")
            action_sell = menu.addAction("SELL")

            action = menu.exec(self.table.viewport().mapToGlobal(pos))

            if action ==action_sell:
                self.table.setItem(row, column, QTableWidgetItem("SELL"))
                self.audit.info(f"MODIFICATION SIDE : col = {column}, row = {row}, val = SELL")
            elif action == action_buy:
                self.table.setItem(row,column, QTableWidgetItem("BUY"))
                self.audit.info(f"MODIFICATION SIDE : col = {column}, row = {row}, val = BUY")

        if column == 5 : 
            
            action_market = menu.addAction('Market')
            action_limit = menu.addAction('Limit')
            action_stop = menu.addAction("Stop")
            action_stoplimit = menu.addAction("StopLimit")

            action = menu.exec(self.table.viewport().mapToGlobal(pos))

            if action == action_limit:
                self.table.setItem(row, column, QTableWidgetItem("Limit"))
                self.audit.info(f"MODIFICATION ORDERTYPE : col = {column}, row = {row}, val = LIMIT")
            elif action == action_market:
                self.table.setItem(row,column, QTableWidgetItem("Market"))
                self.audit.info(f"MODIFICATION ORDERTYPE : col = {column}, row = {row}, val = MARKET")
            elif action == action_stop:
                self.table.setItem(row,column, QTableWidgetItem("Stop"))
                self.audit.info(f"MODIFICATION ORDERTYPE : col = {column}, row = {row}, val = STOP")
            elif action == action_stoplimit:
                self.table.setItem(row,column, QTableWidgetItem("StopLimit"))
                self.audit.info(f"MODIFICATION ORDERTYPE : col = {column}, row = {row}, val = STOPLIMIT")
        
        if column == 8 : 
            
            action_day = menu.addAction('day')
            action_gtd = menu.addAction('gtd')
            action_gtc = menu.addAction("gtc")
            action_AtClose = menu.addAction("atclose")
            action_AtOpen = menu.addAction("atopen")

            action = menu.exec(self.table.viewport().mapToGlobal(pos))

            if action == action_day:
                self.table.setItem(row, column, QTableWidgetItem("day"))
                self.audit.info(f"MODIFICATION VALIDITYTYPE: col = {column}, row = {row}, val = DAY")
            elif action == action_gtd:
                self.table.setItem(row,column, QTableWidgetItem("gtd"))
                self.audit.info(f"MODIFICATION VALIDITYTYPE: col = {column}, row = {row}, val = GTD")
            elif action == action_gtc:
                self.table.setItem(row,column, QTableWidgetItem("gtc"))
                self.audit.info(f"MODIFICATION VALIDITYTYPE: col = {column}, row = {row}, val = GTC")
            elif action == action_AtClose:
                self.table.setItem(row,column, QTableWidgetItem("atclose"))
                self.audit.info(f"MODIFICATION VALIDITYTYPE: col = {column}, row = {row}, val = ATCLOSE")
            elif action == action_AtOpen:
                self.table.setItem(row,column, QTableWidgetItem("atopen"))
                self.audit.info(f"MODIFICATION VALIDITYTYPE: col = {column}, row = {row}, val = ATOPEN")
        
        if column == 10:
            actions = {menu.addAction(acc): acc for acc in self.client_accounts}
            action = menu.exec(self.table.viewport().mapToGlobal(pos))

            if action in actions: 
                selected_account = actions[action]
                self.table.setItem(row,column, QTableWidgetItem(selected_account))
                self.audit.info(f"MODIFICATION ACCOUNT: col = {column}, row = {row}, val = {selected_account}")
        
        if column == 1 :
            if self.no_isin:
                return False 
            isin = self.table.item(row, column - 1 ).text()
            currency_fourni = self.table.item(row, column + 1 ).text()
            dic_exchange_brut = self.info_bloom_dic[isin]
            actions = {}
            list_exchange = []
            
            if dic_exchange_brut : 
                

                for d in dic_exchange_brut.keys():
                    currency = dic_exchange_brut[d]['currency']
                    if currency_fourni in ('Not Found', 'Not Valid', ''):

                        ticker = d
                        exchange = dic_exchange_brut[d]['exchange']
                        if not exchange:
                            exchange = 'Not Valid'
                        if not currency:
                            currency = 'Not Valid'
                        volume_moyen_30 = dic_exchange_brut[d]['volume_moyen']
                        chaine = ticker + ' // ' + exchange + ' // ' + currency.upper() +  ' // ' + volume_moyen_30
                        list_exchange.append(chaine)
                    elif currency_fourni == currency.upper():
                        ticker = d
                        exchange = dic_exchange_brut[d]['exchange']
                        volume_moyen_30 = dic_exchange_brut[d]['volume_moyen']
                        if not exchange:
                            exchange = 'Not Valid'
                        if not currency:
                            currency = 'Not Valid'
                        chaine = ticker + ' // ' + exchange + ' // ' + currency.upper() +  ' // ' + volume_moyen_30
                        list_exchange.append(chaine)
                if isin in self.etf:
                    exchange = self.etf[isin][0]
                    currency = self.etf[isin][1]
                    chaine = ' // ' + exchange + ' // ' + currency.upper() + ' // ' + 'ETF'
                    list_exchange.append(chaine)

                actions = {menu.addAction(acc): acc for acc in list_exchange}
                action = menu.exec(self.table.viewport().mapToGlobal(pos))

            elif isin in self.etf:
                exchange = self.etf[isin][0]
                currency = self.etf[isin][1]
                chaine = ' // ' + exchange + ' // ' + currency.upper() + ' // ' + 'ETF'
                list_exchange.append(chaine)

                actions = {menu.addAction(acc): acc for acc in list_exchange}
                action = menu.exec(self.table.viewport().mapToGlobal(pos))


            if actions != {}:
                if action in actions:
                    selected_account = actions[action]
                    selected_account = selected_account.split('//')
                    exchange = selected_account[1].strip()
                    for idx in range(self.table.rowCount()):
                        if self.table.item(idx, 0).text() == isin:
                            if exchange in self.mapping_bloom:
                                exchange = self.mapping_bloom[exchange]
                            self.table.setItem(idx,column, QTableWidgetItem(exchange))
                            self.table.setItem(idx,column + 1, QTableWidgetItem(selected_account[2].strip()))
                            self.audit.info(f"MODIFICATION EXCHANGE/CURRENCY: row = {idx}, Exchange = {selected_account[1].strip()} / Currency = {selected_account[2].strip()} ")
                        
                            


    # Menu contextuel pour les colonnes
    def show_column_context_menu(self, position):
        """Affiche le menu contextuel pour les colonnes"""
        col_index = self.table.horizontalHeader().logicalIndexAt(position)
        if col_index < 0:
            return
            
        col_name = self.table.horizontalHeaderItem(col_index).text()
        
        context_menu = QMenu(self)

        if col_index == 10: 
            actions = {context_menu.addAction(acc): acc for acc in self.client_accounts}
            action = context_menu.exec(self.table.viewport().mapToGlobal(position))

            if action in actions: 
                selected_account = actions[action]
                for row in range(self.table.rowCount()):
                    self.table.setItem(row, col_index, QTableWidgetItem(selected_account))
        
        elif col_index== 1:
            if self.no_isin :
                return False 
            maximum_volume =QAction("Maximum Volume",self)
            maximum_volume.triggered.connect(lambda: self.maximum_volume_exchange(col_index))
            context_menu.addAction(maximum_volume)

            context_menu.exec(self.table.horizontalHeader().mapToGlobal(position))
                
        
        else:
            # Remplir toute la colonne
            fill_action = QAction(f"Remplir toute la colonne '{col_name}'", self)
            fill_action.triggered.connect(lambda: self.fill_entire_column(col_index))
            context_menu.addAction(fill_action)
            
            # Remplacer dans la colonne
            replace_action = QAction(f"Remplacer dans la colonne '{col_name}'", self)
            replace_action.triggered.connect(lambda: self.replace_in_column(col_index))
            context_menu.addAction(replace_action)
            
            # Vider la colonne
            clear_action = QAction(f"Vider la colonne '{col_name}'", self)
            clear_action.triggered.connect(lambda: self.clear_column(col_index))
            context_menu.addAction(clear_action)
            
            context_menu.exec(self.table.horizontalHeader().mapToGlobal(position))

    # Remplir une colonne entière
    def fill_entire_column(self, col_index):
        """Remplit toute la colonne avec une valeur donnée"""
        col_name = self.table.horizontalHeaderItem(col_index).text()
        
        value, ok = QInputDialog.getText(
            self, 
            "Remplir la colonne", 
            f"Entrez la valeur pour remplir toute la colonne '{col_name}':"
        )
        
        if ok:
            for row in range(self.table.rowCount()):
                item = QTableWidgetItem(str(value))
                
                if value == '':
                    item.setBackground(QColor(255, 255, 200))
                elif value == "Not Found" or value == 'Not Valid':
                    item.setBackground(QColor(255, 150, 150))
                else:
                    item.setBackground(QColor(240, 240, 240))
                
                self.table.setItem(row, col_index, item)
                self.audit.info(f"Fill entire column, col = {col_index}, row = {row}, val = {item}")

    # Remplacer dans une colonne
    def replace_in_column(self, col_index):
        """Remplace toutes les occurrences d'une valeur dans la colonne"""
        col_name = self.table.horizontalHeaderItem(col_index).text()
        
        old_value, ok1 = QInputDialog.getText(
            self, 
            "Remplacer dans la colonne", 
            f"Valeur à remplacer dans la colonne '{col_name}':"
        )
        
        if ok1:
            new_value, ok2 = QInputDialog.getText(
                self, 
                "Remplacer dans la colonne", 
                f"Nouvelle valeur pour remplacer '{old_value}':"
            )
            
            if ok2:
                replacement_count = 0
                
                for row in range(self.table.rowCount()):
                    item = self.table.item(row, col_index)
                    if item and item.text() == old_value:
                        item.setText(str(new_value))
                        self.audit.info(f'Remplacement dans la colonne {col_name} de {old_value} par {new_value}')
                        replacement_count += 1
                        
                        if new_value == '':
                            item.setBackground(QColor(255, 255, 200))
                        elif new_value == "Not Found" or new_value == 'Not Valid':
                            item.setBackground(QColor(255, 150, 150))
                        else:
                            item.setBackground(QColor(240, 240, 240))
                
                QMessageBox.information(
                    self, 
                    "Remplacement terminé", 
                    f"{replacement_count} valeur(s) remplacée(s) dans la colonne '{col_name}'"
                )

    # Vider une colonne
    def clear_column(self, col_index):
        """Vide toute la colonne"""
        col_name = self.table.horizontalHeaderItem(col_index).text()
        
        reply = QMessageBox.question(
            self, 
            "Confirmer", 
            f"Êtes-vous sûr de vouloir vider toute la colonne '{col_name}'?"
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            for row in range(self.table.rowCount()):
                item = QTableWidgetItem("")
                item.setBackground(QColor(255, 255, 200))
                self.table.setItem(row, col_index, item)
    
    def maximum_volume_exchange(self,col_index):
        """Fonction selectionnant l'exchange avec le volume mensuel le plus important pour une currency donnée. Si la currency est absente, la fonction ne fait rien """
        for idx in range(self.table.rowCount()):
                isin = self.table.item(idx, col_index - 1 ).text()
                currency_fourni = self.table.item(idx, col_index + 1 ).text()
                if currency_fourni not in ('Not Found', 'Not Valid', ''):
                    dic_exchange_brut = self.info_bloom_dic[isin]
                    if dic_exchange_brut:
                        liste_couple_exchange_currency = []
                        liste_volume = []
                        for d in dic_exchange_brut.keys():
                            exchange = dic_exchange_brut[d]['exchange']
                            currency_bloom = dic_exchange_brut[d]['currency']
                            if not exchange:
                                exchange = 'Not Valid'
                            if not currency_bloom:
                                currency_bloom = 'Not Valid'
                            
                            if currency_fourni.upper() == currency_bloom.upper():
                                liste_couple_exchange_currency += [[exchange,currency_fourni]]
                                liste_volume.append(int(dic_exchange_brut[d]['volume_moyen'].replace(',','')))
                        if liste_volume != []:
                            index, value = max(enumerate(liste_volume),key  = lambda x : x[1])
                            exchange = liste_couple_exchange_currency[index][0].strip().upper()
                            if exchange in self.mapping_bloom:
                                exchange = self.mapping_bloom[exchange]
                            
                            if exchange == 'NOT VALID':
                                item = QTableWidgetItem(exchange)
                                item.setBackground(QColor(255,83,73))
                                self.table.setItem(idx,col_index, item)
                            else:
                                
                                self.table.setItem(idx,col_index, QTableWidgetItem(exchange))
                            self.audit.info(f"MODIFICATION EXCHANGE/CURRENCY MAXIMUM VOLUME: row = {idx}, Exchange = {exchange} / Currency = {liste_couple_exchange_currency[index][1].strip().upper()} ")



    def _build_dataframe(self):
        """Fonction fabriquant l'interface dataframe"""
        data = {}
        for attr in self.ordered_attrs:
            if hasattr(self.FILE, attr):
                val = getattr(self.FILE, attr)
                if isinstance(val, pd.Series):
                    data[attr] = val.reset_index(drop=True)
        self.data = data
        return pd.DataFrame(data)
    
    def save_correction(self):
        """Creation du fichier csv final"""

        file_depotpath = ''
        if self.use_info.folder_id == 'PROD':
            file_depotpath = PROD_FOLDER
        elif self.use_info.folder_id == 'UAT':
            file_depotpath = UAT_FOLDER
        elif self.use_info.folder_id == 'TEST':
            file_depotpath = TEST_FOLDER


        self.save_button.setEnabled(False)
        new_data = {}

        for col in range(self.table.columnCount()):
            col_name = self.table.horizontalHeaderItem(col).text()
            corrected_values = []
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col)
                corrected_values.append(item.text() if item else None)
            new_data[col_name] = pd.Series(corrected_values)

        for name, series in new_data.items():
            setattr(self.FILE, name, series)
        self.new_data = new_data
        corrected_df = pd.DataFrame(new_data)

        base, ext = os.path.splitext(self.FILE.csv_path)
    
        self.new_path = f"FR_{FILE.clientUlID}_Ordres_{FILE.time}{ext}"
        self.new_path = os.path.join(file_depotpath,self.new_path)
        corrected_df.to_csv(self.new_path, index=False,header=False,lineterminator="\r\n", sep=';', encoding="utf-8")        
        
        os.startfile(self.new_path)
        QMessageBox.information(self, "Succès", f"Corrections sauvegardées et fichier exporté : {self.new_path}")
        
        self.FILE.audit.info("DataFrame sortie : \n%s", corrected_df.to_string(index=False))
        self.FILE.audit.info(f"Corrections sauvegardées et fichier exporté : {self.new_path}")
        self.force = True 
        self.close()

    def closeEvent(self,event):
        """Fonction controlant la fermeture forcé ou non du programme"""
        if not self.force:
            QApplication.closeAllWindows()
            QApplication.quit()
            event.accept()
            self.fclose.close()

class CorrectionButton(QPushButton):
    """Classe de l'interface permettant l'ajout de données dans les tables de mapping du type TBL_BTRADE..."""
    def __init__(self,client_name,text = 'ADD CORRECTIONS', parent = None):
        super().__init__(text,parent)
        self.client_name = client_name
    
    def open_correction(self):
        self.window = MappingApp(audit,['isin','Exchange','Currency','Side','Quantity','OrderType','ValidityType','Account','Price','StopPrice','ExpireDate','Reference','Comment'],self.client_name)
        self.window.show()


# =================================================================================================
# =                                           SECTION: REGLES
# =================================================================================================



# =================================================================================================
# =                                           SECTION: BONUS
# =================================================================================================
 
def excel_to_df(df,cell):
    """Retourne la valeur d'une cellule dans le DF à partir d'une postion excel (ex A1, B2 etc )"""
    col_letter = ''.join([c for c in cell if c.isalpha()]).upper()
    row_number = int(''.join([c for c in cell if c.isdigit()])) - 1 

    col_number = 0 
    for i, letter in enumerate(reversed(col_letter)):
        col_number += (string.ascii_uppercase.index(letter) + 1 )* (26**i)
    col_number -=1
  
    return df.iloc[row_number-1,col_number]

def date_transform(val):
    """
    Verifie si 'val' est une date.
    - Si c'est une date != aujourd'hui : retourne date YYYYMMDD
    - Si c'est la date du jour : retourne day 
    - Sinon : retourne val 
    """
    if not val:
        return val 
    
    #Liste des formats possibles
    date_formats = [
        "%Y-%m-%d", "%d/%m/%Y","%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y","%Y%m%d","%d.%m.%Y"
    ]
   
    for fmt in date_formats:
        try:
            date_obj = datetime.strptime(str(val),fmt).date()
            today =datetime.today().date()
            if date_obj == today :
                
                return 'day'
            else:
                return date_obj.strftime("%Y%m%d")
        except ValueError:
            continue
    return val 



def bloom_format_transformation_ticker(ticker):
    """Retourne le ticker bloom sous la forme "XX XX Equity" """
    if len(ticker.split(" ")) == 2:
        ticker = ticker + " Equity"
        
    return ticker

# =================================================================================================
# =                                           MAIN
# =================================================================================================



if __name__ == "__main__": 
    pd.set_option('display.max_rows',None)
    args = sys.argv[1:]
    file_path = args[0]
    name_client = args[1]
    date = args[2]
    subfolder_path = args[3] 
    fileloggerpath = args[4]
    user_id = args[5]
    folder_id = args[6]
    absolut_path = args[7]

    with open(f'{absolut_path}\\{CONFIG}','r',encoding='utf-8') as f :
        dic = json.load(f)

    TEST_FOLDER = dic['path_depot_test_local']
    UAT_FOLDER = dic['uatfolder_path']
    PROD_FOLDER = dic['prodfolder_path']


    Use_info = UseInfo(user_id=user_id,folder_id=folder_id)
    audit = AuditLogger(fileloggerpath)

    print("Fichier bien reçu :", file_path)

    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    
    fclose = forceclose(audit,subfolder_path)
    FILE = Mapping_program_fonction(file_path,name_client,date,audit,fclose,absolut_path,Use_info)
    window = CorrectionWindow(FILE,fclose)
    window.show()
    
    sys.exit(app.exec())
