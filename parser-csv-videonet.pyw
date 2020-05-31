# Python 3.6
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
from tkinter import Tk, Frame, BOTH
import csv
import os
import re
import datetime


class MyFrame(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.initUI()

    def initUI(self):
        self.pack(fill=BOTH, expand=1)
        self.parent.title("Parser VideoNet CSV")

        w = 480
        h = 220 
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()        
        x = (sw - w)/2
        y = (sh - h)/2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        lbl = Label(self, text='Необходимо выбрать отчет VideoNet в формате "csv"', font=("Helvetica", 12))
        lbl.grid(row=0, column=1, columnspan=5, sticky=W, padx=10, pady=5)
        scroll = Scrollbar(self, orient=HORIZONTAL)
        scroll.grid(row=3, column=2, columnspan=2, sticky=E+W)
        self.output = Entry(self, width=60, justify=LEFT, xscrollcommand=scroll.set)
        self.output.grid(row=2, column=2, columnspan=2, sticky=E+W)
        scroll.config(command=self.output.xview)
        self.button = Button(self, text="Выбрать файл", command=self.load_file, width=12)
        self.button.grid(row=2, column=1, sticky=W, pady=0)
        self.lbl3 = Label(self, text="*Имя файла при\nсохранении: ")
        self.lbl3.grid(row=4, column=1, sticky=W)
        self.obj3 = Def_file_name()        
        self.obj3.namedef()
        self.deftext = StringVar()
        self.deftext.set(self.obj3.name_def)
        self.inputname = Entry(self, textvariable=self.deftext, width=30, fg="gray22")
        self.inputname.grid(row=4, column=2, sticky=W)
 
        self.button4 = Button(self, text="Запустить", command=self.start, width=12)
        self.button4.grid(row=5, column=1, sticky=W, pady=20)
  
    def load_file(self):
        self.selected_file = askopenfilename(initialdir = "./",
                                     title = "Select file",
                                     filetypes = (("CSV files","*.csv"),
                                                  ("all files","*.*")))
        if self.selected_file:
            self.output.delete("0","end")
            self.output.insert("0", self.selected_file)

    def start(self):
        text_result = StringVar()
        file_exist = StringVar()
        lbl_result = Label(self, textvariable = text_result, width=45, anchor=W, justify=LEFT, fg='green', font = ('Sans','10','bold'))
        lbl_result.grid(row=5, column=2, columnspan=2, sticky=W, padx=0)
        lbl_file_exist = Label(self, textvariable = file_exist, width=15, anchor=W, justify=LEFT, fg='red', font=("Sans",'10'))
        lbl_file_exist.grid(row=6, column=1, columnspan=2, sticky=W, padx=2)
        
        MyFrame.name_for_save = self.deftext.get()
        obj = Pars_csv()
        selected_file_final = self.output.get()
        if selected_file_final:
            if os.path.exists(selected_file_final):
                file_exist.set('')
                if MyFrame.name_for_save:
                    obj.processing(selected_file_final)
                    text_result.set('Готово!')
                else:
                    MyFrame.name_for_save = self.obj3.name_def
                    obj.processing(selected_file_final)
                    text_result.set('Готово! Имя файла: ' + MyFrame.name_for_save)
            else:
                file_exist.set('Файл не найден!')
        else:
            file_exist.set('Файл не выбран!')
                  
class Def_file_name():
    def namedef(self):
        self.now = datetime.datetime.now()
        self.now_s = self.now.strftime("%Y-%m-%d-%H-%M") 
        self.name_def = 'VideoNet-'+self.now_s+'.csv'
    

class Pars_csv():
    def writecsv(self, current_date_full, source, entry_or_exit_temp, user):
        with open(MyFrame.name_for_save, 'a', newline='', encoding='cp1251') as csvfile:                        
                csvfile = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                csvfile.writerow([current_date_full, source, entry_or_exit_temp, user])        
    
    def processing(self, selected_file):
        first_event_for_user = []
        last_event_for_user = []
        first_event_for_date = []
        last_event_for_date = []
        first_event_for_source = []
        last_event_for_source = []
        first_entry_or_exit = []
        last_entry_or_exit = []
        sutki = 0
        pattern = r"В[ы]?ход\s{1,}\w{2,}\s{1,}\w{2,}" 
        with open(selected_file, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            row_count = len(list(reader))       
        self.writecsv('Дата и время', 'Источник', 'Событие', 'ФИО')
        with open(selected_file, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            for row in reader:
                if reader.line_num <=3:
                    continue
                else:
                    current_date_full = row[0]
                    source = row[1]
                    cur_text = row[2]
                    current_date = current_date_full[:10]
                    result = re.search(pattern, cur_text)                 
                    if result:
                        user = result[0][5:].lstrip()
                        entry_or_exit_temp = result[0][:5].rstrip()                      
                        if current_date == sutki:
                            if user in first_event_for_user:
                                if user in last_event_for_user:
                                    for index, value in enumerate(last_event_for_user): 
                                        if value == user: 
                                            last_event_for_date[index] = current_date_full
                                            last_event_for_source[index] = source
                                            last_entry_or_exit[index] = entry_or_exit_temp
                                else: 
                                    last_event_for_user.append(user)
                                    last_event_for_date.append(current_date_full)
                                    last_event_for_source.append(source)
                                    last_entry_or_exit.append(entry_or_exit_temp)
                            else: 
                                first_event_for_date.append(current_date_full)
                                first_event_for_source.append(source)
                                first_entry_or_exit.append(entry_or_exit_temp)
                                first_event_for_user.append(user)
                                self.writecsv(current_date_full, source, entry_or_exit_temp, user)                                                      
                        else:
                            sutki = current_date
                            for index_us, value_us in enumerate(last_event_for_user):                            
                                self.writecsv(last_event_for_date[index_us], last_event_for_source[index_us], last_entry_or_exit[index_us], value_us)
                            first_event_for_user.clear()
                            last_event_for_user.clear()
                            first_event_for_date.clear()
                            first_entry_or_exit.clear()
                            first_event_for_source.clear()
                            last_event_for_date.clear()
                            last_entry_or_exit.clear()
                            last_event_for_source.clear()

                            first_event_for_date.append(current_date_full)
                            first_event_for_source.append(source)
                            first_entry_or_exit.append(entry_or_exit_temp)
                            first_event_for_user.append(user)
                            self.writecsv(current_date_full, source, entry_or_exit_temp, user)
                    if reader.line_num == row_count: 
                        for index_us, value_us in enumerate(last_event_for_user): 
                            self.writecsv(last_event_for_date[index_us], last_event_for_source[index_us], last_entry_or_exit[index_us], value_us)

                        
def main():
    root = Tk()
    app = MyFrame(root)
    root.mainloop() 


if __name__ == "__main__":
    main()
