
def get_stylesheet():
    return """
        QWidget {
            background-color:#E5E7EA;
            font-family: 'Segoe UI', 'Helvetica Neue', sans-serif;
        }
        #sidebar {
            background-color:#2C3E50;
            color: white;
            border-right: 2px solid #1A252F;
        }

        #sidebarTitle {
            font-size: 22px;
            font-weight: bold;
            padding: 20px 10px;
            border-bottom: 1px solid #1A252F;
        }

        QPushButton{
            background-color: #34495E;
            color : white;
            border: none;
            padding: 12px
            font-size: 15px;
            text-align: left;
            border-radius: 6px;
        }

        QPushButton:hover{
        background-color: #3E5871;
        }

        QPushButton:pressed {
        background-color: #2E4053;
        }

        #QFrameDataExchange{
            border: 2px solid #4CAF50;
            border-radius: 10px;
            background-color: #f5f5f5;
            padding: 10px;
        }

        #welcomeLabel{
            font-size: 24px;
            color: #2C3E50;
        }
        #labelconfigWindow  {
            color : #2C3E50;
            font-size: 14px;
            font-weight: 500;
            }
        #fieldconfigWindow{
            background-color: white;
            border: 1 px solid #BDC3C7;
            padding: 6px 8px;
            font-size: 14px;
            selection-background-color: #2980B9;
        }
        #fieldconfigWindow:focus {
            border: 2px solid #2980B9
            background-color: #ECF6FB
            }

        #search_barEi{
            background-color: white;
            border: 1 px solid #BDC3C7;
            padding: 6px 8px;
            font-size: 14px;
            selection-background-color: #2980B9;
        }
        #search_barEi:focus {
            border: 2px solid #2980B9
            background-color: #ECF6FB
            }
        
        #list_widgetEI {
            background-color: #FFFFFF;
            border-radius: 6px;
            border: 1px solid #BDC3C7;
            outline: none;
            selection-background-color: 3498DB;
        }

        QWidget#Memory, QWidget#Memory * {
            background-color: #000000;
            font-family: 'Segoe UI', 'Helvetica Neue', sans-serif;
        }

        #list_widgetEI::item{
            padding: 6px 10px;
        }
        #list_widgetEI::item:selected{
            backgroung-color: #3498DB;
            color: white;
            vorder-radius: 4px;}
        
        """
