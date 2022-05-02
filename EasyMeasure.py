import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton, QRadioButton, QLabel, QLineEdit,QFileDialog, QSpinBox,QDialogButtonBox,QFormLayout,QVBoxLayout,QCheckBox,QMessageBox,QComboBox,QFileDialog
from PyQt5.QtGui import QIcon, QFont, QRegExpValidator
from PyQt5.QtCore import QRegExp

from LSMMacroClass import LSMMacroClass # importing class from package 
from analyse import Analyse

class MacroGeneratorForm(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "EasyMeasure"
        self.top = 700
        self.left = 300
        self.width = 500
        self.height = 500
        self.icon = QIcon('lidrotec.jpg')
        self.vertical = 8

        self.macro_generator_details = MacroGeneratorDetails()

        self.validator = QRegExpValidator(QRegExp(r'[0-9]+'))
        

        self.initUI()


    def initUI(self):
        self.setGeometry(self.top, self.left, self.width, self.height)
        self.setWindowTitle(self.title)
        self.setWindowIcon(self.icon)


        self.label_4 = QLabel('Select parameters to measure: ',self)
        self.label_4.move(5,self.vertical+25)
        self.parameter_1 = QCheckBox("depth",self)
        self.parameter_1.setChecked(True)
        self.parameter_1.move(225,self.vertical+22)
        self.parameter_2 = QCheckBox("areal roughness",self)
        self.parameter_2.move(280,self.vertical+22)
        self.parameter_3 = QCheckBox("surface roughness",self)
        self.parameter_3.move(380,self.vertical+22)

        self.label_15 = QLabel('Select template to save the reports: ',self)
        self.label_15.move(5,self.vertical+75)
        self.template = QLineEdit(self)
        self.template.resize(220,20)
        self.template.move(225,self.vertical+72)
        self.push_button_6= QPushButton('...',self)
        self.push_button_6.resize(25,20)
        self.push_button_6.move(450,self.vertical+72)
        self.push_button_6.clicked.connect(lambda: self.browse(self.push_button_6))




        self.label_5 = QLabel('Select alignment file for the measurement: ',self)
        self.label_5.move(5,self.vertical+100)
        self.alignment = QLineEdit(self)
        self.alignment.resize(220,20)
        self.alignment.move(225,self.vertical+97)
        self.push_button_2 = QPushButton('...',self)
        self.push_button_2.resize(25,20)
        self.push_button_2.move(450,self.vertical+97)
        self.push_button_2.clicked.connect(lambda: self.browse(self.push_button_2))


        self.label_6 = QLabel('Number of measurement fields:              row',self)
        self.label_6.move(5,self.vertical+125)
        self.label_7 =QLabel('column',self)
        self.label_7.move(275,self.vertical+125)
        self.row = QSpinBox(self)
        self.row.move(225,self.vertical+122)
        self.column = QSpinBox(self)
        self.column.move(310,self.vertical+122)


        self.label_8 = QLabel('Dimension of the pattern (?):',self)
        self.label_8.move(5,self.vertical+150)
        self.label_8.setToolTip('<img src="pattern_dimension.jpg">')
        self.label_9 = QLabel('size (in µm)',self)
        self.label_9.move(165,self.vertical+147)
        self.measurement_length = QLineEdit(self)
        self.measurement_length.move(225,self.vertical+150)
        self.measurement_length.resize(60, 20)
        self.measurement_length.setValidator(self.validator)
        self.label_10 = QLabel("gap (in µm)",self)
        self.label_10.move(165,self.vertical+174)
        self.gap = QLineEdit(self)
        self.gap.move(225,self.vertical+174)
        self.gap.resize(60, 20)
        self.gap.setValidator(self.validator)

        self.label_12 = QLabel('Select objective for the measurement: ',self)
        self.label_12.move(5,self.vertical+200)
        self.objective = QComboBox(self)
        self.objective.move(225,self.vertical+200)
        self.objective.addItem("5x")
        self.objective.addItem("10x")
        self.objective.addItem("20x")
        self.obj = "5x"
        self.objective.currentTextChanged.connect(self.objective_changed)

        self.label_13 = QLabel('Multi-Area stitching :                                 nX',self)
        self.label_13.move(5,self.vertical+225)
        self.label_14 =QLabel('nY',self)
        self.label_14.move(295,self.vertical+225)
        self.nX = QSpinBox(self)
        self.nX.setValue(1)
        self.nX.move(225,self.vertical+222)
        self.nY = QSpinBox(self)
        self.nY.move(310,self.vertical+222)
        self.nY.setValue(1)

        self.label_11 = QLabel("Save the report as: ",self)
        self.label_11.move(5,self.vertical+275)
        self.extension = QComboBox(self)
        self.extension.move(225,self.vertical+275)
        self.extension.addItem("EXCEL")
        self.extension.addItem("PDF")
        self.ext = "EXCEL"
        self.extension.currentTextChanged.connect(self.ext_changed)


        self.label_2 = QLabel('Name of the report to save the data :',self)
        self.label_2.move (5,self.vertical+300)
        self.folder_name = QLineEdit(self)
        self.folder_name.move(225,self.vertical+297)

        self.label_3 = QLabel('Select directory to save the reports: ',self)
        self.label_3.move(5,self.vertical+325)
        self.folder_directory = QLineEdit(self)
        self.folder_directory.resize(220,20)
        self.folder_directory.move(225,self.vertical+322)
        self.push_button_1 = QPushButton('...',self)
        self.push_button_1.resize(25,20)
        self.push_button_1.move(450,self.vertical+322)
        self.push_button_1.clicked.connect(lambda: self.browse(self.push_button_1))

        self.push_button_4 = QPushButton('Back',self)
        self.push_button_4.move(310,400)
        self.push_button_4.clicked.connect(self.back)

        self.push_button_5 = QPushButton('Continue',self)
        self.push_button_5.move(390,400)
        self.push_button_5.clicked.connect(self.passingInfo)


    def back(self):
        self.previous_window = Window1()
        self.previous_window.show()
        self.hide()
         
    def browse(self,push_button):
        if push_button == self.push_button_1:
            self.directory = os.path.normpath(QFileDialog.getExistingDirectory(self,'Select the directory to save the file'))
            self.folder_directory.setText(self.directory)

        if push_button == self.push_button_2:
            self.directory, filter = QFileDialog.getOpenFileName(self,'Select the alignment file', filter='alignment (*.isc)')
            if self.directory:
                self.directory_ = os.path.normpath(self.directory)
                self.alignment.setText(self.directory_)

        if push_button == self.push_button_6:
            self.directory, filter = QFileDialog.getOpenFileName(self,'Select the template to analyse', filter='Templates (*.tpl)')
            if self.directory:
                self.directory_ = os.path.normpath(self.directory)
                self.template.setText(self.directory_)



    def passingInfo(self):
        self.macro_generator_details.folder_name.setText(self.folder_name.text())
        self.macro_generator_details.folder_name.adjustSize()
        self.macro_generator_details.folder_directory.setText(self.folder_directory.text())
        self.macro_generator_details.folder_directory.adjustSize()

        #array for selected parameters for evaluation
        self.macro_generator_details.selected_parameters = []
        if self.parameter_1.isChecked() == True:
            self.macro_generator_details.selected_parameters.append(self.parameter_1.text())
        if self.parameter_2.isChecked() == True:
            self.macro_generator_details.selected_parameters.append(self.parameter_2.text())
        if self.parameter_3.isChecked() == True:
            self.macro_generator_details.selected_parameters.append(self.parameter_3.text())
        self.macro_generator_details.required_parameters.setText(', '.join(self.macro_generator_details.selected_parameters))

        self.macro_generator_details.required_parameters.adjustSize()
        self.macro_generator_details.template.setText(self.template.text())
        self.macro_generator_details.template.adjustSize()
        self.macro_generator_details.alignment.setText(self.alignment.text())
        self.macro_generator_details.alignment.adjustSize()
        self.macro_generator_details.row.setNum(self.row.value())
        self.macro_generator_details.row.adjustSize()
        self.macro_generator_details.column.setNum(self.column.value())
        self.macro_generator_details.column.adjustSize()
        self.macro_generator_details.measurement_length.setText(self.measurement_length.text())
        self.macro_generator_details.measurement_length.adjustSize()
        self.macro_generator_details.gap.setText(self.gap.text())
        self.macro_generator_details.gap.adjustSize()

        self.macro_generator_details.objective.setText(self.obj)

        self.macro_generator_details.nX.setNum(self.nX.value())
        self.macro_generator_details.nY.setNum(self.nY.value())


        self.macro_generator_details.extension.setText(self.ext)


        self.macro_generator_details.displayInfo()

    def ext_changed(self, ext):
        self.ext = ext
        
    def objective_changed(self,obj):
        self.obj = obj




class MacroGeneratorDetails(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "EasyMeasure"
        self.top = 700
        self.left = 300
        self.width = 500
        self.height = 500 
        self.icon = QIcon('lidrotec.jpg')
        self.vertical = 8

        self.selected_parameters = []

        self.initUI()

    def initUI(self):
        self.setGeometry(self.top, self.left, self.width, self.height)
        self.setWindowTitle(self.title)
        self.setWindowIcon(self.icon)

        self.head_label = QLabel('Please check the details given and click confirm to create the macro',self)
        self.head_label.move(5,self.vertical)
        self.head_label.setFont(QFont("Helvetica [Cronyx]",10))


        self.label_4 = QLabel('Selected parameters for the measurement: ',self)
        self.label_4.move(5,self.vertical+25)
        self.required_parameters = QLabel(self)
        self.required_parameters.move(225,self.vertical+25)

        self.label_14 = QLabel('Selected template for the measurement: ',self)
        self.label_14.move(5,self.vertical+50)
        self.template = QLabel(self)
        self.template.move(225,self.vertical+50)
       

        self.label_5 = QLabel('Selected alignment file for the measurement: ',self)
        self.label_5.move(5,self.vertical+75)
        self.alignment = QLabel(self)
        self.alignment.move(225,self.vertical+75)
       

        self.label_6 = QLabel('Number of measurement fields:              row:',self)
        self.label_6.move(5,self.vertical+100)
        self.label_7 =QLabel('column:',self)
        self.label_7.move(275, self.vertical+100)
        self.row = QLabel(self)
        self.row.move(225,self.vertical+100)
        self.column = QLabel(self)
        self.column.move(315,self.vertical+100)

        self.label_8 = QLabel('Size of the measurement field(in µm):',self)
        self.label_8.move(5,self.vertical+125)
        self.measurement_length = QLabel(self)
        self.measurement_length.move(225,self.vertical+125)

        self.label_9 = QLabel("Gap between the marking(in µm): ",self)
        self.label_9.move(5,self.vertical+150)
        self.gap = QLabel(self)
        self.gap.move(225,self.vertical+150)

        self.label_11 = QLabel('Select objective for the measurement: ',self)
        self.label_11.move(5,self.vertical+175)
        self.objective = QLabel(self)
        self.objective.move(225,self.vertical+175)

        self.label_12 = QLabel('Multi-Area stitching :                                 nX:',self)
        self.label_12.move(5,self.vertical+200)
        self.label_13 =QLabel('nY:',self)
        self.label_13.move(295,self.vertical+200)
        self.nX = QLabel(self)
        self.nX.move(225,self.vertical+200)
        self.nY = QLabel(self)
        self.nY.move(315,self.vertical+200)

        self.label_2 = QLabel('Name of the file to save the report :',self)
        self.label_2.move (5,self.vertical+250)
        self.folder_name = QLabel(self)
        self.folder_name.move(225,self.vertical+250)

        self.label_3 = QLabel('Directory to save the reports: ',self)
        self.label_3.move(5,self.vertical+275)
        self.folder_directory = QLabel(self)
        self.folder_directory.move(225,self.vertical+275)

        self.label_10 = QLabel("The report will be saved as: ",self)
        self.label_10.move(5,self.vertical+300)
        self.extension = QLabel(self)
        self.extension.move(225,self.vertical+300)


        self.push_button_1 = QPushButton('Edit',self)
        self.push_button_1.move(310,400)
        self.push_button_1.clicked.connect(self.back)

        self.push_button_2 = QPushButton('Confirm',self)
        self.push_button_2.move(390,400)
        self.push_button_2.clicked.connect(self.create_macro)

        self.message = QMessageBox()
        self.message.setWindowIcon(self.icon)
        self.message.setWindowTitle(self.title)
        self.message.setText("Your macro is ready!")


    def displayInfo(self):
        self.show()


    def back(self):
        self.hide()



    def create_macro(self):

        self.LSMMC = LSMMacroClass()

        #to save file in user defined directory
        self.LSMMC.macroFileName,self.filter = QFileDialog.getSaveFileName(self, 'Save File',filter='macro (*.mcr)')
        
        self.LSMMC.generateFile()
        self.LSMMC.comSetAutoSaveSetting(self.folder_directory.text(), self.folder_name.text())
        #selecting the template according to the required parameter
        # if self.required_parameters.text() == "depth":
        #     self.LSMMC.comSetReportTemplate(self.template_directory+'\multipleLine_depth_tiltCorrected_range100_minDepth10.tpl')
        # if self.required_parameters.text() == "depth, surface roughness":
        #     self.LSMMC.comSetReportTemplate(self.template_directory+'\multiple_line_depth_roughness.tpl')
        # if self.required_parameters.text() == "depth, areal roughness":
        self.LSMMC.comSetReportTemplate(self.template.text())
        self.LSMMC.comSetAlignment(self.alignment.text())

        #selecting the objective lens
        if self.objective.text()=="5x": 
            self.LSMMC.comChangeRevolver(1)
        if self.objective.text()=="10x": 
            self.LSMMC.comChangeRevolver(2)
        if self.objective.text()=="20x":
            self.LSMMC.comChangeRevolver(3)

        xStart = (int(self.measurement_length.text()))/2
        yStart =-1*(int(self.measurement_length.text()))/2
        repNr = 1
        for iy in range(0,int(self.row.text())*int(self.gap.text()),int(self.gap.text())):
            for ix in range(0,int(self.column.text())*int(self.gap.text()),int(self.gap.text())):
                self.LSMMC.comMoveXYZStage(xStart+ix,yStart-iy)

                self.LSMMC.comSnapshot(1)
                self.LSMMC.comSetMultiPointArea(int(self.nX.text()),int(self.nY.text()))

                # now start your measurements
                self.LSMMC.com3DExtendedAcquisition()

                # save report as PDF-file
                self.LSMMC.comSaveReport(repNr,self.extension.text())
                self.LSMMC.comCloseReport()
                repNr = repNr + 1
        self.LSMMC.closeFile()

        self.message.exec_()
        self.back()




class LsmReportAnalyser(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "EasyMeasure"
        self.top = 700
        self.left = 300
        self.width = 500
        self.height = 500
        self.icon = QIcon('lidrotec.jpg')
        self.vertical = 8

        self.initUI()

    def initUI(self):
        self.setGeometry(self.top, self.left, self.width, self.height)
        self.setWindowTitle(self.title)
        self.setWindowIcon(self.icon)

        self.label_1 = QLabel('Select any of the excel report: ',self)
        self.label_1.move(5,30)
        self.folder_directory = QLineEdit(self)
        self.folder_directory.move(225,25)

        # self.label_2 = QLabel('Name of the excel files: ',self)
        # self.label_2.move(5,self.vertical+25)
        # self.file_name = QLineEdit(self)
        # self.file_name.move(225,self.vertical+25)

        self.push_button_1 = QPushButton('Browse',self)
        self.push_button_1.move(400,25)
        self.push_button_1.clicked.connect(lambda: self.browse(self.push_button_1))

        self.label_5 = QLabel('Select parameters and its excel cell number: ',self)
        self.label_5.move(5,self.vertical+50)
        self.parameter_1 = QCheckBox("depth",self)
        self.parameter_1.setChecked(True)
        self.parameter_1.move(225,self.vertical+47)
        self.parameter_2 = QCheckBox("areal roughness",self)
        self.parameter_2.move(225,self.vertical+72)
        self.parameter_3 = QCheckBox("surface roughness",self)
        self.parameter_3.move(225,self.vertical+97)
        self.depth_cells = QLineEdit(self)
        self.depth_cells.move(340,self.vertical+47)
        self.areal_roughness_cells = QLineEdit(self)
        self.areal_roughness_cells.move(340,self.vertical+72)  
        self.surface_roughness_cells = QLineEdit(self)
        self.surface_roughness_cells.move(340,self.vertical+97)          

        self.label_3 = QLabel('Number of measurement fields:              row',self)
        self.label_3.move(5,self.vertical+125)
        self.label_4 =QLabel('column',self)
        self.label_4.move(275,self.vertical+125)
        self.row = QSpinBox(self)
        self.row.move(225,self.vertical+122)
        self.column = QSpinBox(self)
        self.column.move(310,self.vertical+122)
        
        self.push_button_4 = QPushButton('Back',self)
        self.push_button_4.move(310,300)
        self.push_button_4.clicked.connect(self.back)

        self.push_button_5 = QPushButton('Create',self)
        self.push_button_5.move(390,300)
        self.push_button_5.clicked.connect(self.create_report)

        #to show message after creating the report
        self.message = QMessageBox()
        self.message.setWindowIcon(self.icon)
        self.message.setWindowTitle(self.title)
        self.message.setText("The final report is ready!")

       


    def browse(self,push_button):
        if push_button.text() == 'Browse':
            self.directory, filter = QFileDialog.getOpenFileName(self,'Select any one of the excel report generated by LSM', filter='Excel (*.xlsx)')
            if self.directory:
                self.directory_ = os.path.normpath(self.directory)
                self.folder_directory.setText(self.directory_)

    def back(self):
        self.previous_window = Window1()
        self.previous_window.show()
        self.hide()

    def create_report(self):
        self.analyse = Analyse(self.folder_directory.text())

        #array for selected parameters for evaluation
        if self.parameter_1.isChecked() == True:
            self.analyse.selected_parameters.append(self.parameter_1.text())
            self.analyse.depth_cell_number = self.depth_cells.text()
        if self.parameter_2.isChecked() == True:
            self.analyse.selected_parameters.append(self.parameter_2.text())
            self.analyse.areal_roughness_cell_number = self.areal_roughness_cells.text()
        if self.parameter_3.isChecked() == True:
            self.analyse.selected_parameters.append(self.parameter_3.text())
            self.analyse.surface_roughness_cell_number = self.surface_roughness_cells.text()
        print(self.analyse.selected_parameters[0])
    
        self.analyse.create_master_sheet(int(self.row.text()),int(self.column.text()))
        self.message.exec_()
        self.back()





class Window1(QWidget):
    
    def __init__(self):

        super(QWidget,self).__init__()
        
        # window specifications
        self.title = "EasyMeasure"
        self.top = 700
        self.left = 300
        self.width = 500
        self.height = 500
        self.icon = QIcon('lidrotec.jpg')

        #font properties 
        self.instruction_font = QFont("Courier",12)
        self.label_font = QFont("Courier",10)


        self.initUI()



    def initUI(self):

        self.setGeometry(self.top, self.left, self.width, self.height)
        self.setWindowTitle(self.title)
        self.setWindowIcon(self.icon)


        # self.big_layout = QVBoxLayout()
        # self.small_layout = QVBoxLayout()


        self.pushButton = QPushButton('Continue',self)
        self.pushButton.move(390,300)
        

        self.label = QLabel('Hello! What would you like to do now ?',self)
        self.label.move(75,100)
        self.label.setFont(self.instruction_font)
        self.radioButton1 = QRadioButton('Generate macro for LSM',self)
        self.radioButton1.setFont(self.label_font)
        self.radioButton1.move(160,150)
        self.radioButton1.toggled.connect(lambda:self.radio_button_state(self.radioButton1))
        self.radioButton2 = QRadioButton('LSM Data Analysis', self)
        self.radioButton2.setFont(self.label_font)
        self.radioButton2.move(160,175)
        self.radioButton2.toggled.connect(lambda:self.radio_button_state(self.radioButton2))

        # self.small_layout.addWidget(self.radioButton1)
        # self.small_layout.addWidget(self.radioButton2)

        # self.big_layout.addWidget(self.label)
        # self.big_layout.addLayout(self.small_layout)
        # self.big_layout.addWidget(self.pushButton)

        # self.setLayout(self.big_layout)




        self.show()

    # This method uses the state of the radio button and push button ('continue') and move to required window 
    def radio_button_state(self,radio_button):

        if radio_button.text() == 'Generate macro for LSM':
            if radio_button.isChecked() == True:
                self.pushButton.clicked.connect(self.macro_generator)

        if radio_button.text() == 'LSM Data Analysis':
            if radio_button.isChecked() == True:
                self.pushButton.clicked.connect(self.lsm_report_analyse)


    def macro_generator(self):                             
        self.w = MacroGeneratorForm()
        self.w.show()
        self.hide()

    def lsm_report_analyse(self):
        self.w = LsmReportAnalyser()
        self.w.show()
        self.hide()


def main():

    app = QApplication(sys.argv)

    intro = Window1()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()