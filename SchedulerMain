import tkinter as tk
from tkinter import *
import os, threading
import xlwings as xw
import pandas as pd
import io, schedule, time
from pathlib import Path
from datetime import datetime

ListToRun = []
temp_path = ""
temp_unit = ""
temp_time = ""
temp_at = ""
logstr = ""
temp_logstr = " "
LargeFont = ("Verdana", 12)
SmallFont = ("Verdana",8)


# Call to load the init file to
class ItemX:
   def __init__(self,Num,xlsm,Module,Sub,TimeX,ScheUnitX, AtX, optX): 
       self.Num = Num
       self.xlsm = xlsm
       self.Module = Module
       self.Sub = Sub
       self.TimeX = TimeX 
       self.ScheUnitX = ScheUnitX
       self.AtX = AtX
       self.optX = optX

class mainframe(Frame):

    def __init__(self):
        super().__init__()
        self.main_page()

    def main_page(self):
        global temp_logstr
        self.master.title("VBA Scheduler")
        self.pack(fill=BOTH, expand=True)
        self.columnconfigure(1, weight=2)
        self.columnconfigure(2, weight=2)
        self.columnconfigure(3, weight=4)
        self.columnconfigure(4, weight=0)
        self.columnconfigure(5, weight=0)
        self.rowconfigure(3, weight=1)
        self.rowconfigure(5, pad=3)

        self.loaded_config_path = StringVar()
        self.loaded_config_time = StringVar()
        self.loaded_config_unit = StringVar()
        self.loaded_config_at = StringVar()
        self.logstr =StringVar()

        self.logstr.set(temp_logstr)

        self.activity_log = Label(self, text= self.logstr, font = LargeFont, borderwidth=1, relief="groove", justify=CENTER, anchor="w")
        self.activity_log.grid(row=1, column=0, columnspan=2, rowspan=6,pady=5,padx=5, sticky=E+W+S+N)

        temp_path, temp_time, temp_unit, temp_at = LoadInit()
        self.loaded_config_path.set(temp_path)
        self.loaded_config_time.set(temp_time)
        self.loaded_config_unit.set(temp_unit)
        self.loaded_config_at.set(temp_at)

        self.loadedpath = Label(self,textvariable= self.loaded_config_path, font = LargeFont, borderwidth=0, relief="groove", justify=RIGHT, anchor="w")
        self.loadedpath.grid(row=3, column=3, columnspan=2, rowspan=2,pady=5,padx=2, sticky=E+W+S+N)

        self.loadedtime = Label(self,textvariable= self.loaded_config_time, font = LargeFont, borderwidth=0, relief="groove", justify=CENTER, anchor="w")
        self.loadedtime.grid(row=3, column=4, columnspan=1, rowspan=2,pady=5,padx=0, sticky=E+W+S+N)

        self.loadedunit = Label(self,textvariable= self.loaded_config_unit, font = LargeFont, borderwidth=0, relief="groove", justify=CENTER, anchor="w")
        self.loadedunit.grid(row=3, column=5, columnspan=1, rowspan=2,pady=5,padx=0, sticky=E+W+S+N)

        self.loadedat = Label(self,textvariable= self.loaded_config_at, font = LargeFont, borderwidth=0, relief="groove", justify=LEFT, anchor="w")
        self.loadedat.grid(row=3, column=6, columnspan=2, rowspan=2,pady=5,padx=5, sticky=E+W+S+N)

        self.button_reload = Button(self, text="Reload", font = LargeFont, height =1, width = 15,command= self.updatelabel)
        self.button_reload.grid(row=1, column=3)

        self.button_run = Button(self, text="Run", font = LargeFont, height =1, width = 15,command=run_loaded_schedule_thread)
        self.button_run.grid(row=1, column=4, pady=5, padx=5)

        self.button_stop = Button(self, text="Stop", font = LargeFont, height =1, width = 15)
        self.button_stop.grid(row=1, column=5, pady=5, padx=5)

        self.button_manual = Button(self, text="Run manually", font = LargeFont, height =1, width = 15, command=add_manual)
        self.button_manual.grid(row=6, column=4, pady=5, padx=5)

        self.button_config = Button(self, text="Schedule setting", font = LargeFont, height =1, width = 15, command= add_setting_schedule)
        self.button_config.grid(row=6, column=3, pady=5, padx=5)

    def updatelabel(self):
        temp_path, temp_time, temp_unit, temp_at = LoadInit()
        self.loaded_config_path.set(temp_path)
        self.loaded_config_time.set(temp_time)
        self.loaded_config_unit.set(temp_unit)
        self.loaded_config_at.set(temp_at)
        
    def updatelog(self):
        self.logstr.set(temp_logstr)

def manual_page():
    child_manual = tk.Toplevel()
    child_manual.geometry("700x400")
    child_manual.master.title("Run a VBA manually")
    child_manual.columnconfigure(1, weight=1)
    child_manual.columnconfigure(3, pad=7)
    child_manual.rowconfigure(3, weight=1)
    child_manual.rowconfigure(5, pad=7)
    button_run_manual = tk.Button(child_manual, text="Run", font = LargeFont, height =1, width = 15)
    button_run_manual.grid(row=1, column=4, pady=5, padx=5)

def setting_schedule_page():
    child_setting = tk.Toplevel()
    child_setting.geometry("700x400")
    child_setting.master.title("Schedule setting")
    child_setting.columnconfigure(1, weight=1)
    child_setting.columnconfigure(3, pad=7)
    child_setting.rowconfigure(3, weight=1)
    child_setting.rowconfigure(5, pad=7)
    button_run_manual = tk.Button(child_setting, text="Run", font = LargeFont, height =1, width = 15)
    button_run_manual.grid(row=1, column=4, pady=5, padx=5)

def LoadInit():
    source_path = Path(__file__).resolve()
    source_dir = source_path.parent
    df_CSVinit = pd.read_csv(str(source_dir)+'\Init.csv', header=0)
    df_CSVinit.dropna(how='all',inplace=True)
    df_CSVinit.dropna(axis=1,inplace=True)  #int(df_CSVinit.loc[xRow,"LSGW_SN"]
    temp = df_CSVinit.to_string(columns=['Path To Xlsm', 'Time Scheduled','Unit Scheduled','At'],header=True,index=False, justify="center")
    temp_path = df_CSVinit.to_string(columns=['Path To Xlsm'],header=True,index=False, justify="center")
    temp_time = df_CSVinit.to_string(columns=['Time Scheduled'],header=True,index=False, justify="center")
    temp_unit = df_CSVinit.to_string(columns=['Unit Scheduled'],header=True,index=False, justify="center")
    temp_at = df_CSVinit.to_string(columns=['At'],header=True,index=False, justify="center")

        #print(temp)
    for xRow in df_CSVinit.index:
        #print(df_CSVinit.loc[xRow,"LSGW_SN"])
        ListToRun.append(
            ItemX(
                xRow,
                str(df_CSVinit.loc[xRow,'Path To Xlsm']),
                str(df_CSVinit.loc[xRow,'Module Name']),
                str(df_CSVinit.loc[xRow,'Sub Name']),
                str(df_CSVinit.loc[xRow,'Time Scheduled']),
                str(df_CSVinit.loc[xRow,'Unit Scheduled']),
                str(df_CSVinit.loc[xRow,'At']),
                str(df_CSVinit.loc[xRow,'option']),
            )
        )
    return temp_path, temp_time, temp_unit, temp_at

def RunVBA(xlsmX,ModuleX_SubX,unitx,dayd, opt):
    global temp_logstr
    if unitx == 'month'and int(datetime.now().strftime('%d'))== int(dayd) or  unitx != 'month' :
    #RunIt = RunVBA.RunX(VBAxlsm,VBAMod)
        strToPrint =  str(xlsmX)
        L = strToPrint.split('\\')
        print( 'Start  '+ L[-1]+ '  At  '+ str(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        temp_logstr = temp_logstr + 'Start  '+ L[-1]+ '  At  '+ str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))+ "\n"
       # mainframe.updatelog()
        #temp_logstr =  temp_logstr + "\n"

        timestart = datetime.now()
        xl_app = xw.App(visible=False, add_book=False)
        wb = xl_app.books.open(xlsmX)
        app_pid = xw.apps.keys()

        if unitx == 'month' and int(opt) == 1:
            info_sheet = xw.apps(app_pid[0]).books(L[-1]).sheets('Info')
            info_sheet.range('F6').value = 1
            info_sheet.range('F7').value = 1
            info_sheet.range('F8').value = 1
            info_sheet.range('F17').value = 0
        elif  unitx != 'month' and int(opt) == 1:
            info_sheet = xw.apps(app_pid[0]).books(L[-1]).sheets('Info')
            info_sheet.range('F6').value = 1
            info_sheet.range('F7').value = 1
            info_sheet.range('F8').value = 0
            info_sheet.range('F17').value = 0

        wb.app.calculation = 'manual'
        wb.app.screen_updating = False
        wb.app.display_alerts = False
        run_macro = wb.app.macro(ModuleX_SubX)
        run_macro()

        if len(xw.apps)==1:
            wb.app.calculation = 'automatic' 
            wb.app.screen_updating = True
            wb.app.display_alerts = False
            xl_app.kill()
        elif len(xw.apps)> 1:
            wb.app.calculation = 'automatic'
            wb.app.screen_updating = True
            wb.app.display_alerts = False
            wb.save()
            wb.close
            xl_app.quit()
            #print(len(xw.apps))

        elapsedtime = datetime.now()-timestart
        temp_logstr =  temp_logstr + 'End    '+ 'At  '+ str(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) +'   Elapsed time {}'.format(elapsedtime) + "\n"
        print( 'End    '+ 'At  '+ str(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) +'   Elapsed time {}'.format(elapsedtime) )
        temp_logstr =  temp_logstr + "\n"
        #mainframe.updatelog()
        print(' ')
        print(temp_logstr)
    return

def goaround():
    os.system("start C:/Users/vincent.gouraud/Desktop/") #C:\Users\vincent.gouraud\Desktop\Gainz Always Gainzzzz\Scheduled Run VBA py

def add_manual():
    runmanualwindow = manual_page()
 
def add_setting_schedule():
    runmanualwindow =setting_schedule_page()

def run_loaded_schedule_thread():
    t1 = threading.Thread(target=run_loaded_schedule)
    t1.daemon = True
    t1.start()
    


def run_loaded_schedule():
    print('started something')  
    for X in ListToRun:
        if X.TimeX.isdigit() :
            if X.ScheUnitX == 'minute'  :
                if X.AtX == 'x':
                    if  int(X.TimeX) == 1 :
                        schedule.every().minute.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                    else:
                        schedule.every(int(X.TimeX)).minutes.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().minute.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.ScheUnitX== 'hour' :
                if X.AtX == 'x':
                    if  int(X.TimeX) == 1 :
                        schedule.every().hour.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                    else:
                        schedule.every(int(X.TimeX)).hours.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().hour.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))

            elif X.ScheUnitX == 'day' :
                if X.AtX == 'x':
                    if  int(X.TimeX) == 1 :
                        schedule.every().day.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                    else:
                        schedule.every(int(X.TimeX)).days.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().day.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))

            elif X.ScheUnitX == 'week' :
                if X.AtX == 'x':
                    if  int(X.TimeX) == 1 :
                        schedule.every().week.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                    else:
                        schedule.every(int(X.TimeX)).weeks.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.ScheUnitX == 'month' :
                if X.AtX != 'x':
                    #schedule.every(31).days.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX)) 
                    schedule.every().day.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                    #else:
                    #schedule.every(int(X.TimeX)).weeks.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
        else:
            if X.TimeX == 'monday':
                if X.AtX != 'x':
                    schedule.every().monday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().monday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'tuesday':
                if X.AtX != 'x':
                    schedule.every().tuesday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().tuesday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'wednesday':
                if X.AtX != 'x':
                    schedule.every().wednesday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().wednesday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'thursday':
                if X.AtX != 'x':
                    schedule.every().thursday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().thursday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'friday':
                if X.AtX != 'x':
                    schedule.every().friday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().friday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'saturday':
                if X.AtX != 'x':
                    schedule.every().saturday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().saturday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
            elif X.TimeX == 'sunday':
                if X.AtX != 'x':
                    schedule.every().sunday.at(X.AtX).do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
                else:
                    schedule.every().sunday.do(RunVBA,*(X.xlsm,str(X.Module + '.' + X.Sub),X.ScheUnitX,X.TimeX,X.optX))
    while True:
        schedule.run_pending()
        time.sleep(1)

def main():
    root = tk.Tk()
    root.geometry('1300x500')
    root.title("VBA Scheduler")
    #loaded_config = tk.StringVar()
    app = mainframe()

    root.mainloop()

if __name__ == '__main__':
    main()
    #threading.Thread(target=main).start()

