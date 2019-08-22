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
        h = 250 
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()        
        x = (sw - w)/2
        y = (sh - h)/2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        lbl = Label(self, text='Необходимо выбрать отчет VideoNet в формате "csv"', font=("Helvetica", 12))
        lbl.grid(row=0, column=1, columnspan=3, padx=10, pady=5)

        scroll = Scrollbar(self, orient=HORIZONTAL)
        scroll.grid(row=3, column=2, columnspan=2, sticky=E+W)
        self.output = Entry(self, width=60, justify=LEFT, xscrollcommand=scroll.set)
        self.output.grid(row=2, column=2, columnspan=2, sticky=W)
        scroll.config(command=self.output.xview)
        self.button = Button(self, text="Выбрать файл", command=self.load_file, width=12)
        self.button.grid(row=2, column=1, sticky=W, pady=0)
        #self.button.place(x=150, y=50)
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
        self.lbl4 = Label(self, text="* Не обязательный параметр.")
        self.lbl4.grid(row=6, column=1, sticky=W, columnspan=2, pady=20)

        #self.button5 = Button(self, text="Сохранить как..", command=self.save_file, width=14)
        #self.button5.grid(row=6, column=2, sticky=W)
        #quitButton.place(x=50, y=50)
  
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
        lbl_result = Label(self, textvariable = text_result, width=55, anchor=W, justify=LEFT)
        lbl_result.grid(row=5, column=2, columnspan=2, sticky=W, padx=0)
        self.name_for_save = self.deftext.get()
        obj = Pars_csv()
        selected_file_final = self.output.get()
        if self.name_for_save:
            obj.processing(self.name_for_save, selected_file_final)
            text_result.set('Готово! Имя файла: ' + self.name_for_save)
        else:
            obj.processing(self.obj3.name_def, selected_file_final)
            text_result.set('Готово! Имя файла (по умолчанию): ' + self.obj3.name_def)
                    
class Def_file_name():
    def namedef(self):
        self.now = datetime.datetime.now()
        self.now_s = self.now.strftime("%Y-%m-%d-%H-%M") 
        self.name_def = 'VideoNet-'+self.now_s+'.csv'
   
class Pars_csv():  
    def processing(self, name_for_save, selected_file):
        arr1 = []
        arr2 = []
        period = []
        users = []
        users_out = []
        date_s = []
        door = []
        sutki = 0
        current_user = 0
        n = 0
        pattern = r"Вход\s{1}\w{2,}\s{1}\w{2,}" # Шаблон "Вход Фамилия"
        pattern2 = r"Выход\s{1}\w{2,}\s{1}\w{2,}" # Шаблон "Выход Фамилия"
##        if os.path.isfile(name_for_save):
##            os.remove(name_for_save)
##        else:
##            print("Error: %s file not found" % name_for_save)
            
        with open(name_for_save, 'w', newline='', encoding='cp1251') as csvfile:
            csvfile = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csvfile.writerow(['Дата и время', 'Источник', 'Событие'])
        with open(selected_file, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            for row in reader:
                n += 1
                if n <=3:
                    continue
                else:
                    cur_date_full = row[0]
                    cur_date = cur_date_full
                    cur_source = row[1]
                    cur_text = row[2]
                    arr1.extend(cur_date)
                    arr2.extend(cur_source)
                    cur_date = cur_date[:2]
                    count = len(period)
                    current_count = count - 1
                    past_count = count - 2
                    result = re.search(pattern, cur_text)# Поиск по шаблону "Вход Фамилия"
                    result2 = re.findall(pattern2, cur_text)# Поиск по шаблону "Выход Фамилия"

                    # Разбиваем данные на сутки
                    if sutki == 0:
                        period.append(cur_date)
                        sutki = cur_date
                    elif current_count >= 0 and cur_date != period[current_count]:
                        period.append(cur_date)
                        sutki = cur_date

                        for index_us, value_us in enumerate(users_out):
                            with open(name_for_save, 'a', newline='', encoding='cp1251') as csvfile:                        
                                csvfile = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                                csvfile.writerow([date_s[index_us], door[index_us], value_us])                                          
                        users.clear()
                        users_out.clear()
                        date_s.clear()

                    if result:
                        user = result[0]
                        if user in users:
                            continue
                        else:
                            users.append(user)                            
                            csvtemp = cur_date + cur_source + cur_text
                            with open(name_for_save, 'a', newline='', encoding='cp1251') as csvfile:
                                csvfile = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                                csvfile.writerow([cur_date_full, cur_source, user])
    ##                                writer = csv.writer(f)
    ##                                writer.writerow(csvtemp)
    ##                                f.close()
                    if result2:
                        user_out = result2[0]
                        if user_out in users_out:
                            for index, value in enumerate(users_out):
                                if value == user_out:
                                    #users_out[index] = user_out
                                    date_s[index] = cur_date_full
                                    door[index] = cur_source
                        else:
                            users_out.append(user_out)
                            date_s.append(cur_date_full)
                            door.append(cur_source)
                            #print(user_out + cur_date_full)


def main():
    root = Tk()
    app = MyFrame(root)
    root.mainloop() 


if __name__ == "__main__":
    main()
