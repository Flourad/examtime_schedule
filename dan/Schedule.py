#-*- coding: utf-8 -*-
import colo_group
from PyQt4 import QtGui
from PyQt4 import QtCore
from xlwt import Workbook
import sys
import xlwt

buttons = []
x_list = []
y_list = []
xy_list = []
dif_x = []
dif_y = []
rowMax_list = []
teacher_exam = {}
all_table = {}
pos_examInfom = {}
block_stu = {}
exam_stu1 = {}
exam_stu2 = {}
mapping  = {}  
examBlock_layout = QtGui.QVBoxLayout()
examTime_layout  = QtGui.QGridLayout()

book = Workbook()
sheet1 = book.add_sheet('Sheet1',cell_overwrite_ok = True)


GREEN_TABLE_HEADER = xlwt.easyxf('font: bold 1, name Tahoma, height 160;'
                                'align: vertical center, horizontal center, wrap on;'
                                'borders: left thin, right thin, top thin, bottom thin;'
                                'pattern: pattern solid, pattern_fore_colour 3, pattern_back_colour 3'
                                )

BLUE_TABLE_HEADER = xlwt.easyxf('font: bold 1, name Tahoma, height 160;'
                                'align: vertical center, horizontal center, wrap on;'
                                'borders: left thin, right thin, top thin, bottom thin;'
                                'pattern: pattern solid, pattern_fore_colour 44, pattern_back_colour 44'
                                )

class Button(QtGui.QLineEdit):
    def __init__(self, title, parent):
        super(Button, self).__init__(title, parent)
        self.leftButton = False
    
    def mouseMoveEvent(self, e):
        if e.buttons() != QtCore.Qt.LeftButton:
            return
        self.leftButton = True
        mimeData = QtCore.QMimeData()
        mimeData.setText('%d,%d' % (e.x(),e.y())) 
        pixmap = QtGui.QPixmap.grabWidget(self)

        painter = QtGui.QPainter(pixmap)
        painter.setCompositionMode(painter.CompositionMode_DestinationIn)
        painter.fillRect(pixmap.rect(), QtGui.QColor(0, 0, 0, 127))
        painter.end()
    
        drag = QtGui.QDrag(self)
        drag.setMimeData(mimeData)
        drag.setPixmap(pixmap)
        drag.setHotSpot(e.pos())

        if drag.exec_(QtCore.Qt.CopyAction | QtCore.Qt.MoveAction) == QtCore.Qt.MoveAction:
            print 'moved'
        else:
            print 'copied'
            
    def mousePressEvent(self, e):
        QtGui.QLineEdit.mousePressEvent(self, e)
        if e.button() == QtCore.Qt.RightButton:
            print 'press'

        
class Exam_Schedule(QtGui.QWidget):
    
    def __init__(self):
        super(Exam_Schedule, self).__init__()
        self.leftButton = False
        self.scrollingWidget = QtGui.QWidget()
        self.initUI()
    
    def initUI(self):
    
        mainlayout = QtGui.QVBoxLayout()
        firstline_layout = QtGui.QHBoxLayout() 
        secondline_layout = QtGui.QHBoxLayout()  
        thirdline_layout = QtGui.QHBoxLayout()
        scollWidget_layout = QtGui.QVBoxLayout()
    
        headline = QtGui.QLineEdit()
        headline.setFixedWidth(500)
        headline_font = QtGui.QFont("Time", 20, QtGui.QFont.Bold, True)
        headline.setFont(headline_font)
        headline.setFrame(False)
        headline.setText("Examems session de septembre 2012")
        headline_color = headline.palette()
        headline_color.setColor(headline_color.Text, QtGui.QColor(255, 0, 0))
        headline_color.setColor(headline_color.Base, QtGui.QColor(245, 245, 245))
        headline.setPalette(headline_color)
        
        left_infom = QtGui.QLineEdit("ISIMA ISI2 and ISI3")
        left_infom_color = left_infom.palette()
        left_infom_color.setColor(left_infom_color.Text, QtGui.QColor(100, 149, 237))
        left_infom_color.setColor(left_infom_color.Base, QtGui.QColor(245, 245, 245))
        left_infom.setPalette(left_infom_color)
        left_infom.setFrame(False)
        left_infom.setFixedWidth(230)
        left_infom_font = QtGui.QFont("Time", 15, QtGui.QFont.Bold, True)
        left_infom.setFont(left_infom_font)
        
        right_infom = QtGui.QLineEdit("du Lundi 10 sept au jeudi 20 sept 201")
        right_infom.setFixedWidth(430)
        right_infom_color = right_infom.palette()
        right_infom_color.setColor(right_infom_color.Text,QtGui.QColor(100, 149, 237))
        right_infom_color.setColor(right_infom_color.Base, QtGui.QColor(245, 245, 245))
        right_infom_font = QtGui.QFont("Time", 15, QtGui.QFont.Bold, True)
        right_infom.setFont(right_infom_font)
        right_infom.setPalette(right_infom_color)
        right_infom.setFrame(False)
        
        firstline_layout.addWidget(left_infom, 0, QtCore.Qt.AlignLeft)
        firstline_layout.addWidget(headline, 1, QtCore.Qt.AlignCenter)
        firstline_layout.addWidget(right_infom, 0, QtCore.Qt.AlignRight)
            
        label1 = QtGui.QLabel("Le planning des 2e annees est en bleu")
        label1_font = QtGui.QFont("Time", 15, QtGui.QFont.Bold, True)
        label1.setFont(label1_font)
        
        colorLine1 = QtGui.QLineEdit("")
        colorLine1.setFixedWidth(100)
        colorLine1.setFrame(False)
        color1 = colorLine1.palette()
        color1.setColor(color1.Base, QtGui.QColor(84,255,159))
        colorLine1.setPalette(color1)
        
        label2 = QtGui.QLabel("Celui des 3e annees est en vert!!")
        label2_font = QtGui.QFont("Time", 15, QtGui.QFont.Bold, True)
        label2.setFont(label2_font)
        
        colorLine2 = QtGui.QLineEdit("")
        colorLine2.setFixedWidth(100)
        colorLine2.setFrame(False)
        color2 = colorLine2.palette()
        color2.setColor(color2.Base, QtGui.QColor(152, 245, 255))
        colorLine2.setPalette(color2)
        
        secondline_layout.addWidget(label1, 0, QtCore.Qt.AlignLeft)
        secondline_layout.addWidget(colorLine1, 1, QtCore.Qt.AlignCenter)
        secondline_layout.addWidget(label2, 2, QtCore.Qt.AlignCenter)
        secondline_layout.addWidget(colorLine2)
        secondline_layout.setSpacing(1)
        
        scollWidget_layout.addLayout(examBlock_layout)    
    
        self.scrollingWidget = QtGui.QWidget()
        self.scrollingWidget.setLayout(scollWidget_layout)
        
        myScrollArea = QtGui.QScrollArea()
        myScrollArea.setWidgetResizable(True)
        myScrollArea.setEnabled(True)       
        myScrollArea.setMaximumSize(2000, 500)  # optional
        myScrollArea.setWidget(self.scrollingWidget)
        
        def on_slider_moved(value):
            print "new slider position:%i"% value
            
        myScrollArea.connect(myScrollArea, QtCore.SIGNAL("sliderMoved(int"), on_slider_moved)
           
        line1_buttonLayout = QtGui.QHBoxLayout()
        line2_buttonLayout = QtGui.QHBoxLayout()
        
        button1 = QtGui.QPushButton("Modify")
        button2 = QtGui.QPushButton("Save")
        button3 = QtGui.QPushButton("Query")
        button4 = QtGui.QPushButton("Export")
        button5 = QtGui.QPushButton("Import")
        
        thirdline_layout.setAlignment(QtCore.Qt.AlignLeft)
        thirdline_layout.addWidget(button5)
          
        line1_buttonLayout.addStretch(1)
        line1_buttonLayout.addWidget(button1)
        line1_buttonLayout.addWidget(button2)
        
        line2_buttonLayout.addStretch(1)
        line2_buttonLayout.addWidget(button3)
        line2_buttonLayout.addWidget(button4)
        
        mainlayout.addLayout(firstline_layout)
        mainlayout.addLayout(secondline_layout)
        mainlayout.addLayout(thirdline_layout)
        mainlayout.addWidget(myScrollArea)  
        mainlayout.addLayout(line1_buttonLayout)
        mainlayout.addLayout(line2_buttonLayout) 
        self.setLayout(mainlayout)
       
        
        def change():
            reply1 = QtGui.QMessageBox.question(self, 'PyQt', 'Are you sure to change the planning?', QtGui.QMessageBox.Yes,QtGui.QMessageBox.Cancel)
            if reply1 == QtGui.QMessageBox.Yes:
                self.setAcceptDrops(True)
                
        def save():
            reply2 = QtGui.QMessageBox.question(self, 'PyQt', 'Are you sure to save the change?', QtGui.QMessageBox.Yes,QtGui.QMessageBox.Cancel)
            if reply2 == QtGui.QMessageBox.Yes:
                self.setAcceptDrops(False)
                 
        def showDialog():
            text,ok = QtGui.QInputDialog.getText(self,'inputDialog', 'Enter your name:')
            if ok:
                self.w = query_infom(teacher_exam,text)
                self.w.show()
        def export():
            self.export = export_planning()
            self.export.show()
        
        def import_excelFile():
            self.import_excel = import_excel()
            self.import_excel.show()
            
     
        button1.clicked.connect(change)
        button2.clicked.connect(save)
        button3.clicked.connect(showDialog)
        button4.clicked.connect(export)
        button5.clicked.connect(import_excelFile)
        
    def exam_Planning(self,info_table,filename1,filename2,timeAdd_exam1,timeAdd_exam2): 
          
        def time_map(time):
            real_time = ' '
            if time == 2:
                real_time = '10h~12h'
            elif time==4:
                real_time = '13h30~15h30'
            elif time == 5:
                real_time = '16h~18h'
            return real_time 
        
        def exam_planning_ui():
            
            examTime1_layout = QtGui.QGridLayout()
            examTime2_layout = QtGui.QGridLayout()
            examTime3_layout = QtGui.QGridLayout()
            examTime4_layout = QtGui.QGridLayout()
            space1_layout = QtGui.QGridLayout()
            space2_layout = QtGui.QGridLayout()
            
            t1 = QtGui.QLabel('8h')
            t3 = QtGui.QLabel('10h')
            t4 = QtGui.QLabel('11h')
            t5 = QtGui.QLabel('12h')
            t6 = QtGui.QLabel('13h30')
            t7 = QtGui.QLabel('15h30')
            t8 = QtGui.QLabel('16h')
            t9 = QtGui.QLabel('18h')
            
            
            examTime1_layout.setHorizontalSpacing(80)
            examTime1_layout.addWidget( QtGui.QLabel(' '),0,0)
            examTime1_layout.addWidget(t1,0,1)
            examTime1_layout.addWidget( QtGui.QLabel(' '),0,2)
            
            examTime2_layout.setHorizontalSpacing(118)
            examTime2_layout.addWidget(t3,0,0)
            examTime2_layout.addWidget(t4,0,1)
            examTime2_layout.addWidget(t5,0,2)
            
            examTime3_layout.setHorizontalSpacing(230)
            examTime3_layout.addWidget(t6,0,0)
            examTime3_layout.addWidget(t7,0,1)
            
            examTime4_layout.setHorizontalSpacing(250)
            examTime4_layout.addWidget(t8,0,0)
            examTime4_layout.addWidget(t9,0,1)
            
            space1_layout.setHorizontalSpacing(208)
            space1_layout.addLayout(examTime1_layout,0,0) 
            space1_layout.addLayout(examTime2_layout,0,1) 
            
            space2_layout.setHorizontalSpacing(10)
            space2_layout.addLayout(examTime3_layout,0,0) 
            space2_layout.addLayout(examTime4_layout,0,1) 
            
            examTime_layout.setHorizontalSpacing(35)
            examTime_layout.addLayout(space1_layout,0,0)
            examTime_layout.addLayout(space2_layout,0,1,)
            
               
            eachday_grid_list = []
            
            for i in range(0,8):
                grid = QtGui.QGridLayout()
                eachday_grid_list.append(grid)
            
            planning_writeExcel(eachday_grid_list)
            
            top_line = QtGui.QLabel()
            top_line.setFrameStyle(1)
            top_line.setFixedHeight(1)
            
            line_list = []
            for i in range(0,8):
                bottom_line = QtGui.QLabel()
                bottom_line.setFrameStyle(1)
                bottom_line.setFixedHeight(1)
                line_list.append(bottom_line)
    
            examBlock_layout.addLayout(examTime_layout)
            examBlock_layout.addWidget(top_line)           
            for i in range(0,8):
                examBlock_layout.addLayout(eachday_grid_list[i]) 
                examBlock_layout.addWidget(line_list[i])
                
        def planning_writeExcel(eachday_grid_list):  
            
            col = 0
            sheet_col =0 
            day_num = 1
            day_row = 0
            smallBlock_row = 0
            bigBlock_row = 0
            buttons_num = 0
            temp_row_list = [0, 0, 0]
                
            for i in range(2,11):
                if i%2 == 0:
                    if i != 6:
                        sheet1.col(i).width = 0x0d00 + 12000
                        
            sheet1.write(0,2,"8h~10h")           
            sheet1.write(0,4,"10h~12h")
            sheet1.write(0,8,"13h30~15h30")
            sheet1.write(0,10,"16h~18h")
            
    
            for k, v in info_table.items():
                col = col%6
                sheet_col = sheet_col % 12
                is_three_exist = False
                
                if col == 0: 
                    day_row = max(temp_row_list)+bigBlock_row
                    for i in range(2):
                        eachday_grid_list[day_num-1].addWidget(QtGui.QLabel(''),day_row,col)
                        sheet1.write(day_row,col," ")
                        day_row = day_row+1
                    
                    day = QtGui.QLineEdit('%r'%day_num)
                    day.setFrame(False)
                    day.setFixedWidth(80)
                    eachday_grid_list[day_num-1].addWidget(day,day_row,col)
                    sheet1.write(day_row,col,day_num)
                        
                    day_num += 1
                    col = col + 1
                    sheet_col = sheet_col+2
                        
                    rowMax_list.append(max(temp_row_list))
                    bigBlock_row = max(temp_row_list)+bigBlock_row                   
                                    
                    bigBlock_row =bigBlock_row+1  
                    temp_row_list = []
                        
                if col == 1:
                    blank1 = QtGui.QLabel('')
                    blank1.setFixedWidth(300)
                    eachday_grid_list[day_num-2].addWidget(blank1,bigBlock_row,col)
                    col=col+1
                    sheet_col = sheet_col+2
                        
                if col == 3:
                    blank2 = QtGui.QLabel(' ')
                    blank2.setFixedWidth(20)
                    eachday_grid_list[day_num-2].addWidget(blank2,bigBlock_row,col)
                    col=col+1
                    sheet_col = sheet_col+2  
                                       
                smallBlock_row = bigBlock_row    
                temp_row = 0
                
                three_hour_exam = ''
                
                for s in v[0]:
                    if s[0] in timeAdd_exam1+timeAdd_exam2:
                        is_three_exist = True
                        three_hour_exam = s[0]
                        break
                      
                for s in v[0]:   
                    buttons.append(Button(','.join(s),self))
                    p=buttons[buttons_num].palette()  
                    p.setColor(p.Base,QtGui.QColor(84,255,159))
                    buttons[buttons_num].setPalette(p)
                    eachday_grid_list[day_num-2].addWidget(buttons[buttons_num],smallBlock_row,col)  
                    buttons[buttons_num].setReadOnly(True)         
                    sheet1.write(smallBlock_row,sheet_col,','.join(s),GREEN_TABLE_HEADER) 
                    time = time_map(col)
                    if is_three_exist:
                        if three_hour_exam == s[0]:
                            buttons[buttons_num].setFixedWidth(400) 
                            time = "13h30~16h30" 
                        else:  
                            buttons[buttons_num].setFixedWidth(290)  
                    
                    teacher_exam[s[3]].append([s[0],str(day_num-1),time,'1'])
                            
                    smallBlock_row += 1
                    buttons_num += 1  
                    temp_row += 1
                                            
                for s in v[1]:
                    buttons.append(Button(','.join(s),self)) 
                    p=buttons[buttons_num].palette()  
                    p.setColor(p.Base,QtGui.QColor(152, 245, 255))
                    buttons[buttons_num].setPalette(p)
                    eachday_grid_list[day_num-2].addWidget(buttons[buttons_num],smallBlock_row,col) 
                    buttons[buttons_num].setReadOnly(True)  
                    sheet1.write(smallBlock_row,sheet_col,','.join(s),BLUE_TABLE_HEADER)
                    time = time_map(col)
                
                    if is_three_exist:
                        if three_hour_exam == s[0]:
                            buttons[buttons_num].setFixedWidth(400)
                            time = "13h30~16h30" 
                            
                        else:
                            buttons[buttons_num].setFixedWidth(290)  
                    teacher_exam[s[3]].append([s[0],str(day_num-1),time,'2'])
                            
                    smallBlock_row += 1
                    buttons_num += 1   
                    temp_row += 1  
                    
                col=col+1  
                sheet_col = sheet_col + 2 
                temp_row_list.append(temp_row)  
                
                if is_three_exist: 
                    blank3 = QtGui.QLabel(' ')
                    blank3.setFixedWidth(185)
                    eachday_grid_list[day_num-2].addWidget(blank3,bigBlock_row,col)
                    col=col+1
                    sheet_col = sheet_col+2  
                         
            if col!=0:
                rowMax_list.append(max(temp_row_list))
                
            return examTime_layout
        
        exam_planning_ui() 
        
    def dragEnterEvent(self, e):  
        e.accept()
            
    def dropEvent(self, e):  
        mime = e.mimeData().text()
        x, y = map(int, mime.split(','))
        
        global pos_examInfom, xy_list,exam_stu1, exam_stu2,all_table,dif_x,dif_y,buttons,rowMax_list,block_stu 
        block_pos = []
        Dicblock_pos = {}
        pos_Infom = {}
        blockRow_list =[]
        block_stu = {}
        start_x = 0
        start_y = 0
        startPos_content = []
        
        for button in buttons:
            x_list.append(button.pos().x())
            y_list.append(button.pos().y())
            xy_list.append('*'.join([str(button.pos().x()),str(button.pos().y())]))
        dif_x = list(set(x_list))
        dif_y = list(set(y_list))
        dif_x.sort()
        dif_y.sort()
    
        for i in range(1,len(rowMax_list)+1):
            blockRow_list.append(sum(rowMax_list[0:i]))
            
        block_num = len(all_table.keys())
        
        for num in range(1,block_num+1):
            block_stu[num] = []
            
        i = 0
        j = 1
        for k,v in all_table.items():
            for s in v[0]:
                flag = '1'
                s0 = s[:] 
                s0.insert(0,flag)
                pos_Infom[xy_list[i]] =s0[:]
                pos_examInfom[xy_list[i]] = [s[0],s[3],exam_stu1[s[0]]]
                block_stu[j].extend(exam_stu1[s[0]])
                i = i+1
            for s in v[1]:
                flag = '2'
                s0 = s[:]
                s0.insert(0,flag)
                pos_Infom[xy_list[i]] = s0[:]
                pos_examInfom[xy_list[i]] = [s[0],s[3],exam_stu2[s[0]]]
                block_stu[j].extend(exam_stu2[s[0]])
                i = i+1   
            block_pos.append((x_list[i-1],dif_y[blockRow_list[(j-1)/3 + 1]-1]))
            Dicblock_pos['*'.join([str(x_list[i-1]),str(dif_y[blockRow_list[(j-1)/3 + 1]-1])])]= j 
            j = j + 1   
        
        for button in buttons:
            if button.leftButton == True:
                start_x = button.pos().x()
                start_y = button.pos().y()
                startPos_content = pos_Infom['*'.join([str(button.pos().x()),str(button.pos().y())])]
                infom = pos_examInfom['*'.join([str(button.pos().x()),str(button.pos().y())])]
                button.leftButton = False
                
        def calculate_drop_pos(dif_pos,dropMouse_pos):
            final_pos = 0
            print dif_pos
            print enumerate(dif_pos)
            for i,m in enumerate(dif_pos):
                if i == len(dif_pos)-1:
                    final_pos = dif_pos[len(dif_pos)-1]
                elif m<= dropMouse_pos < dif_pos[i+1]:
                    final_pos = m
                    break
                else:
                    pass
            print final_pos
            return final_pos
            
        def calculate_block_y(pos_y):
            for i,k in enumerate(block_pos):
                if  pos_y<=block_pos[0][1]:
                    block_y = block_pos[0][1]
                else:
                    if block_pos[i][1]<pos_y<=block_pos[i+1][1]:
                        block_y = block_pos[i+1][1]
            return block_y
            
        def time_map(time):
            real_time = ' '
            if time == 2:
                real_time = '10h~12h'
            elif time==4:
                real_time = '13h30~15h30'
            elif time == 5:
                real_time = '16h~18h'
            return real_time  
                   
        def excel_map(x,y,block):
            excel_x = 1
            excel_y = 1
            for i,k in enumerate(dif_y):
                if y == k:
                    if i==0:
                        excel_x = i+1
                    else:
                        excel_x = i+1+(block-1)/3 
            if x==dif_x[0]:
                excel_y = 4
            elif x == dif_x[1]:
                excel_y = 8
            elif x == dif_x[2]:
                excel_y = 10
            return excel_x,excel_y
            
        def row_time(row_x):
            time_num = 1
            if row_x == dif_x[0]:
                time_num = 2
            if row_x == dif_x[1]:
                time_num = 4
            if row_x == dif_x[2]:
                time_num = 5
            return time_num
            
        def move(final_x, final_y):
            e.source().move(QtCore.QPoint(final_x, final_y))
            e.setDropAction(QtCore.Qt.MoveAction)
                
        def update_query():
            final_time = row_time(final_x)
            before_time = row_time(start_x)
            before = [infom[0],str((start_block_Content-1)/3+1),time_map(before_time)]
            after= [infom[0],str((final_block_Content-1)/3+1),time_map(final_time)]
            examInfom = teacher_exam[infom[1]]
           
            for i,v in enumerate(examInfom):
                flag = v.pop()
                if v == before:
                    after.append(flag)
                    examInfom[i] = after
                    print examInfom[i]
            
            teacher_exam[infom[1]] = examInfom 
           
        def update_excel():
            excel_x0,excel_y0 = excel_map(start_x,start_y,start_block_Content)
            excel_x,excel_y = excel_map(final_x,final_y,final_block_Content)
            sheet1.write(excel_x0,excel_y0,'')
            if startPos_content[0]=='1':
                del startPos_content[0]
                sheet1.write(excel_x,excel_y,startPos_content,GREEN_TABLE_HEADER)  
            else:
                del startPos_content[0]
                sheet1.write(excel_x,excel_y,startPos_content,BLUE_TABLE_HEADER)       
                      
        if e.keyboardModifiers() & QtCore.Qt.ShiftModifier: 
            for button in buttons:
                position = self.scrollingWidget.mapFrom(self, e.pos())
                button.move(position-QtCore.QPoint(x, y))
                self.buttons.append(button)
                e.setDropAction(QtCore.Qt.CopyAction)
                
        else: 
            position = self.scrollingWidget.mapFrom(self, e.pos())
            dropMouse_x = position.x()
            dropMouse_y = position.y()
            final_x = calculate_drop_pos(dif_x, dropMouse_x)
            final_y = calculate_drop_pos(dif_y, dropMouse_y)
                      
            start_block_x = start_x       
            final_block_x = final_x
            start_block_y = calculate_block_y(start_y)  
            final_block_y = calculate_block_y(final_y)  
            start_block_Content = Dicblock_pos['*'.join([str(start_block_x), str(start_block_y)])]
            final_block_Content = Dicblock_pos['*'.join([str(final_block_x), str(final_block_y)])]
                                    
            if not list(set(infom[2]) & set(block_stu[final_block_Content])):
                move(final_x, final_y)
                update_query()
                update_excel()  
            else:
                QtGui.QMessageBox.about(self,'PyQt','The move is not allowed,may conflict!!')
                     
        e.accept() 
    
   
        
             
              
class query_infom(QtGui.QWidget):   
    def __init__(self,teacher_exam, text): 
        super(query_infom,self).__init__()
        self.resize(500, 400)
        self.setWindowTitle("the exam information of "+text)
        table = QtGui.QTableView()
        model = QtGui.QStandardItemModel()
        model.setHorizontalHeaderLabels((u'Course', u'day', u'time'))
        table.setModel(model)
        table.setColumnWidth(0,250)
        table.setEditTriggers(QtGui.QAbstractItemView.NoEditTriggers)
        for k,v in teacher_exam.items():     
                if text in k:
                    for s in v:
                        model.appendRow((QtGui.QStandardItem(s[0]),QtGui.QStandardItem(s[1]),QtGui.QStandardItem(s[2]),))
        table_layout = QtGui.QVBoxLayout()
        table_layout.addWidget(table)
        self.setLayout(table_layout)
       
        
class export_planning(QtGui.QWidget):  
    def __init__(self):
        super(export_planning,self).__init__()
        self.setFixedSize(600,100)
        self.setWindowTitle("Export")
        Location = QtGui.QLabel("Location:")
        line1 = QtGui.QLineEdit()
        
        def selectDirectory():
            self.filename = QtGui.QFileDialog.getSaveFileName(self, "savefile",'.',self.tr("Save Files(*.xls)"))
            line1.setText(self.filename)
         
        def saveFile():
            book.save(self.filename)
            QtGui.QMessageBox.about(self,'PyQt','The file have been saved!')
            
        browseButton = QtGui.QPushButton("Browse...")
        browseButton.clicked.connect(selectDirectory)
        sureButton = QtGui.QPushButton("OK")
        sureButton.clicked.connect(saveFile)
        
        button = QtGui.QHBoxLayout()
        button.addWidget(Location)
        button.addWidget(line1)
        button.addWidget(browseButton)
        button.addWidget(sureButton)
        self.setLayout(button)

class import_excel(QtGui.QWidget):  
    def __init__(self):
        super(import_excel,self).__init__()
        self.setFixedSize(600,200)
        self.setWindowTitle("Import")
        
        global filename1,filename2
        def selectDirectory1():
            global filename1
            filename1 = QtGui.QFileDialog.getOpenFileName(self,"Open file dialog","/","Excel files(*.xls)")
            line1.setText(filename1)
    
        def selectDirectory2():
            global filename2
            filename2 = QtGui.QFileDialog.getOpenFileName(self,"Open file dialog","/","Excel files(*.xls)")
            line2.setText(filename2)

        def saveFile1():
            
            global filename1,filename2,exam_stu1, exam_stu2,all_table,teacher_exam,block_stu 

            def generate_fullMapping(infom1,infom2,mapping):
                table = {}
                len1 = len(infom1.keys())
                len2 = len(infom2.keys())
                if len1 >= len2:
                    for color in infom1.keys():
                        if color in mapping:
                            table[color] = [infom1[color], infom2[mapping[color]]]
                        else:
                            table[color] = [infom1[color], []]
                else:
                    temp_mapping = {}
                    for k, v in mapping.items():
                        temp_mapping[v] = k
                    mapping = temp_mapping.copy()
                    for color in infom2.keys():
                        if color in mapping:
                            table[color] = [infom1[mapping[color]], infom2[color]]
                            
                        else:
                            table[color] = [[], infom2[color]]
                return table
        
            infom1, infom2, mapping, exam_stu1, exam_stu2,timeAdd_exam1,timeAdd_exam2 = colo_group.main(filename1,filename2)
                
            
            all_table = generate_fullMapping(infom1,infom2,mapping)
            
            
            for m,n in all_table.items():
                for s in n[0]:
                    if s not in teacher_exam.keys():
                        teacher_exam[s[3]] = []
                for s in n[1]:
                    if s not in teacher_exam.keys():
                        teacher_exam[s[3]] = []  
                        
            self.aa = Exam_Schedule()            
            self.aa.exam_Planning(all_table,filename1,filename2,timeAdd_exam1,timeAdd_exam2)    
            
            def test(self): 
                book1 = Workbook()
                sheet1 = book1.add_sheet('Sheet1',cell_overwrite_ok = True) 
                sheet1.write(0,0,"Students' name")
                sheet1.write(0,1,"exam")
                sheet1.write(0,2,"day")
                sheet1.write(0,3,"time")   
                
                student_exams1 = {}
                student_exams2 = {}               
                
                for m in exam_stu1.values():
                    for s in m:
                        student_exams1[s] = []
                
               
                

                for k, v in teacher_exam.items():
                    for e in v:
                        print e
                    print 
                    
                for k,v in  exam_stu1.items():
                    for n in v:
                        for s in teacher_exam.values():
                            for m in s:
                                if m[0] == k and m[3]=='1' :
                                    if m not in student_exams1[n]:
                                        student_exams1[n].append(m)
                                        
                i = 1
                j = 1
                for k,v in student_exams1.items():
                    #print k,":"
                    j = i
                    sheet1.write(j,0,k)
                    for m in v:
                        sheet1.write(i,1,m[0])
                        sheet1.write(i,2,m[1])
                        sheet1.write(i,3,m[2])
                        sheet1.write(i,4,m[3])
                        i = i+1
                        #print ','.join(m)
                    #print 
                    
                for m in exam_stu2.values():
                    for s in m:
                        student_exams2[s] = []
                
               
                for k,v in  exam_stu2.items():
                    for n in v:
                        for s in teacher_exam.values():
                            for m in s:
                                if m[0] == k and m[3]=='2' :
                                    if m not in student_exams2[n]:
                                        student_exams2[n].append(m)      
                
                for k,v in student_exams2.items():
                    #print k,":"
                    j = i
                    sheet1.write(j,0,k)
                    for m in v:
                        sheet1.write(i,1,m[0])
                        sheet1.write(i,2,m[1])
                        sheet1.write(i,3,m[2])
                        sheet1.write(i,4,m[3])
                        i = i+1
                        #print ','.join(m)
                    #print 
                
                book1.save("C:\list.xls")
                
                def judge_examtimetable_fairness(student_exams1,student_exams2):
                    is_conflict = False
                    for k,v in student_exams1.items():
                        list = [','.join([s[1],s[2]]) for s in v]    
                        list1 =  [exam[0] for exam in v] 
                        if len(set(list)) != len(list) or len(set(list1)) != len(list1):      
                            print "The plan is conflict!"
                            is_conflict = True
                            
                    for k,v in student_exams2.items():
                        list = [','.join([s[1],s[2]]) for s in v]     
                        list1 =  [exam[0] for exam in v] 
                        if len(set(list)) != len(list) or len(set(list1)) != len(list1):      
                            print "The plan is conflict!"
                            is_conflict = True
                    if not is_conflict:
                        print "考试时间表是可行的！"
                
                judge_examtimetable_fairness(student_exams1,student_exams2)
                         
                            
                
            test(self)
            QtGui.QMessageBox.about(self,'PyQt','The file have been imported!')
            
            
        Location1 = QtGui.QLabel("Location:")
        line1 = QtGui.QLineEdit()
        Location2 = QtGui.QLabel("Location:")
        line2 = QtGui.QLineEdit()  
           
        browseButton1 = QtGui.QPushButton("Browse...")
        browseButton1.clicked.connect(selectDirectory1)
        
        browseButton2 = QtGui.QPushButton("Browse...")
        browseButton2.clicked.connect(selectDirectory2)
        
        sureButton = QtGui.QPushButton("OK")
        sureButton.clicked.connect(saveFile1)
        
        layout1 = QtGui.QHBoxLayout()
        layout1.addWidget(Location1)
        layout1.addWidget(line1)
        layout1.addWidget(browseButton1)
        
        layout2 = QtGui.QHBoxLayout()
        layout2.addWidget(Location2)
        layout2.addWidget(line2)
        layout2.addWidget(browseButton2)
        
        layout3 = QtGui.QHBoxLayout()
        layout3.addStretch(1)
        layout3.addWidget(sureButton)
        
        mylayout = QtGui.QVBoxLayout()
        mylayout.addLayout(layout1)
        mylayout.addLayout(layout2) 
        mylayout.addLayout(layout3)
            
        self.setLayout(mylayout)   

                  
def main():
    app = QtGui.QApplication(sys.argv)
    ex = Exam_Schedule()
    ex.show()        
    app.exec_()
    
if __name__ == '__main__':
    main()