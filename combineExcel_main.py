from  PyQt6.QtWidgets import QApplication,QDialog,QMessageBox,QFileDialog
from  ui_combineExcel import Ui_Dialog_combineExcel
import excelmergeclass
import sys
import os
import threading
import class_private_singnal
import time
import Enum_box


class combine_main(Ui_Dialog_combineExcel,QDialog):


    def __init__(self) -> None:
        super().__init__()
        self.updateui = class_private_singnal.Mysingal()
        self.msgboxType = Enum_box.MsgboxType
     
        self.setupUi(self)
        self.show()

        

        self.pushButton_selectdir.clicked.connect(self.click_selectdir)
        self.pushButton_remove.clicked.connect(self.remove_filename)
        self.pushButton_combine.clicked.connect(self.merge_excels)
        
        #创建自定义信号，可跨线程调用QMessageBox
        self.updateui.setResult.connect(self.update_info)
        self.updateui.setmessagebox.connect(self.update_msgbox)
        self.updateui.setprogressbar.connect(self.update_progressbar)




    
    def update_info(self,result):
        self.lineEdit_dirshow.setText(result)

    def update_msgbox(self,boxType:Enum_box.MsgboxType,t:str,content:str):
        if boxType.value == self.msgboxType.info.value:
            QMessageBox.information(self,t,content)
        if boxType.value == self.msgboxType.warning.value:
            QMessageBox.warning(self,t,content)

    def update_progressbar(self,p:int):
        self.progressBar.setValue(p)

    def click_selectdir(self):
        filedialog = QFileDialog.getExistingDirectory(self,"select dir",os.getcwd())

        #取消选择
        if not filedialog:
            return
        self.updateui.setResult.emit(filedialog)
        

        excel_paths = []

        for excel_path in os.listdir(filedialog):
            if excel_path.startswith("~$"):
                continue

            if excel_path.endswith((".xls","xlsx")):
                excel_paths.append(os.path.join(filedialog,excel_path))

        self.listWidget_filename.clear()    
        self.listWidget_filename.addItems(excel_paths)

    def remove_filename(self):
        selectedItem = self.listWidget_filename.currentIndex().row()

        self.listWidget_filename.takeItem(selectedItem)

    def merge_excels(self):

        def inner_method():
            count = self.listWidget_filename.count()
            if not count:       
                self.updateui.setmessagebox.emit(self.msgboxType.warning,"提示","没有可以合并的excel表格")
                return
            
            path_list =[]
            for idx in range(count):
                path_list.append(self.listWidget_filename.item(idx).text())

            output_path,type = QFileDialog.getSaveFileName(self,"save path",os.getcwd(),filter="Excel files(*.xlsx);;Excel files(*.xls)")
           
           
            #取消保存
            if not output_path:
                return

            self.save_excel(path_list,output_path)

        return inner_method()

    def save_excel(self,path_list,output_path):
        def inner_method():
            self.updateui.setprogressbar.emit(0)
            excelmergeclass.merge_excels(path_list,output_path)
            self.updateui.setprogressbar.emit(100)
            self.updateui.setmessagebox.emit(self.msgboxType.warning,"信息提示","save succeed")

        threading.Thread(target=inner_method).start()
        
            
        

    
    


if __name__ =="__main__":
    app =QApplication(sys.argv)
    
    icombine_main = combine_main()
    sys.exit(app.exec())
