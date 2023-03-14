# -Begin-----------------------------------------------------------------
import datetime
import inspect
import json
import math
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog as fd

import center_tk_window
import openpyxl
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows

# - Functions------------------------------------------------------------
import SAP_Functions


# Check to see if all checkboxes are checked
def check_status():
    if check_v_1.get() & check_v_2.get() & check_v_3.get():
        b1.config(state=tk.NORMAL)
        os.system(f'taskkill /im saplogon.exe /f')
    else:
        b1.config(state=tk.DISABLED)


#####################################

def open_sap(in_queue, ):
    while True:
        item = in_queue.get()
        # process
        SAP_Functions.sap_login(environment)
        in_queue.task_done()


def excel_save(in_queue):
    workbook = openpyxl.load_workbook(project_file)
    sheet = workbook['Result']
    while True:
        item = in_queue.get()
        # process
        query = json.loads(item)
        from datetime import date
        today = str(date.today())
        current_time = str(datetime.datetime.now().strftime("%H:%M:%S"))

        df = pandas.DataFrame([[query["ticket_number_"], today, current_time, query["status"], query["message"], query["scrap_component_"], query["scrap_quantity_"], query["scrap_area_"],
                                query["scrap_sub_area_"], query["scrap_code_"], query["scrap_header_"], query["scrap_cost_center_"]]],
                              columns=["No.Ticket", "Fecha", "Hora", "Estatus", "Resultado_SAP", "Componente",
                                       "Cantidad", "Area", "Subarea", "Codigo_Scrap", "Header", "Centro_Costos"])
        for row_df in dataframe_to_rows(df, header=False, index=False):
            if not row_df[0] is None:
                sheet.append(row_df)
        workbook.save(project_file)
        workbook.close()


def process_sap():
    try:
        global files_
        global label_2
        global textbox
        global sap_instances
        global all_threads
        global stop_threads
        from pathlib import Path

        top = tk.Toplevel(width=275, height=400, cursor="watch")
        top.iconbitmap('./img/image.ico')
        top.title("Movimientos SCRAP Con Formato")
        root.withdraw()
        # center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=20, pady=20)
        label_1 = ttk.Label(top, text="En Proceso")
        label_1.grid(row=4, column=0)

        progress_bar = ttk.Progressbar(top, orient="horizontal", mode="determinate", maximum=100, value=0)
        progress_bar.grid(row=7, column=0)
        progress_bar.start()
        progress_bar.step(10)
        textbox = tk.Text(top, height=9, width=30)
        textbox.grid(row=10, column=0)
        label_2 = ttk.Label(top, )
        label_2.grid(row=5, column=0)

        # SAP_Functions.sap_login(environment)
        thread_excel = threading.Thread(target=excel_save, args=(excel_queue,))
        thread_excel.daemon = True
        thread_excel.start()

        def capture(ticket_number_, message, scrap_component_, scrap_quantity_, scrap_area_, scrap_sub_area_, scrap_code_, scrap_header_, scrap_cost_center_):
            excel_queue.put(json.dumps({"ticket_number_": f"{ticket_number_}", "status": "OK", "message": f"{message}",  "scrap_component_": f"{scrap_component_}", "scrap_quantity_": f"{scrap_quantity_}","scrap_area_": f"{scrap_area_}", "scrap_sub_area_": f"{scrap_sub_area_}", "scrap_code_": f"{scrap_code_}", "scrap_header_": f"{scrap_header_}", "scrap_cost_center_": f"{scrap_cost_center_}"}))
            textbox.insert('1.0', "SAP: " + scrap_component_ + ' Status: ' + "OK\n")

        def err(ticket_number_, message, scrap_component_, scrap_quantity_, scrap_area_, scrap_sub_area_, scrap_code_, scrap_header_, scrap_cost_center_):
            excel_queue.put(json.dumps({"ticket_number_": f"{ticket_number_}", "status": "ERROR", "message": f"{message}", "scrap_component_": f"{scrap_component_}", "scrap_quantity_": f"{scrap_quantity_}",
                                        "scrap_area_": f"{scrap_area_}", "scrap_sub_area_": f"{scrap_sub_area_}", "scrap_code_": f"{scrap_code_}", "scrap_header_": f"{scrap_header_}", "scrap_cost_center_": f"{scrap_cost_center_}"}))
            textbox.insert('1.0', "SAP: " + scrap_component_ + ' Status: ' + "ERR\n")

        def do_work(in_queue):
            global current_process
            global label_2
            current_process = 0
            current_t = threading.current_thread()
            while True:
                item = in_queue.get()
                # process
                query = json.loads(item)
                if current_process == 0:
                    current_process += 1
                progress_bar.destroy()
                label_2.config(text=f"{current_process-1} de {total_r}")
                label_2.grid(row=5, column=0)

                mb1a_response = json.loads(SAP_Functions.mb1a_(query["scrap_material"], query["scrap_header"], query["scrap_code"], storage_location, query["scrap_cost_center"], scrap_order, query["scrap_component"], query["scrap_quantity"],  int(current_t.getName())))
                if mb1a_response["error"] != "N/A":
                    err(query["ticket_number"], mb1a_response["error"], query["scrap_component"], query["scrap_quantity"], query["scrap_area"], query["scrap_sub_area"], query["scrap_code"], query["scrap_header"], query["scrap_cost_center"])
                    current_process += 1
                    in_queue.task_done()
                else:
                    capture(query["ticket_number"], mb1a_response["result"], query["scrap_component"], query["scrap_quantity"], query["scrap_area"], query["scrap_sub_area"], query["scrap_code"], query["scrap_header"], query["scrap_cost_center"])
                    current_process += 1
                    in_queue.task_done()

        for x in range(int(sap_instances)):
            thread_sap = threading.Thread(target=open_sap, name=str(x), args=(sap_queue,))
            thread_sap.daemon = True
            thread_sap.start()

        for z in range(int(sap_instances)):
            sap_queue.put(z)
        sap_queue.join()
        total_r = 0
        for file_ in files_:
            csv_file = pandas.read_csv(file_)
            total_rows = ++csv_file.shape[0]
            total_r = total_r + total_rows
            # No.Ticket # Fecha # Turno # No.Parte # Ensamble # No.Parte # Componenete # Cantidad # Area # Subarea # Codigo # Scrap # Cliente # Header
            for index, row in csv_file.iterrows():
                try:
                    ticket_number = row["ID"]
                    scrap_date = row["DATE"]
                    scrap_shift = row["SHIFT"]
                    scrap_material = row["PART_NUMBER"]
                    scrap_component = row["COMPONENT_NUMBER"]
                    scrap_quantity = row["QUANTITY"]
                    scrap_area = row["AREA"]
                    scrap_sub_area = row["SUBAREA"]
                    scrap_code = row["SCRAP_CODE"]
                    scrap_customer = row["CLIENT"]
                    scrap_header = row["HEADER"]
                    scrap_cost_center = row["COST_CENTER"]

                    work.put(json.dumps(
                        {"ticket_number": f"{ticket_number}",
                         "scrap_date": f"{scrap_date}",
                         "scrap_shift": f"{scrap_shift}",
                         "scrap_material": f"{scrap_material}",
                         "scrap_component": f"{scrap_component}",
                         "scrap_quantity": f"{scrap_quantity}",
                         "scrap_area": f"{scrap_area}",
                         "scrap_sub_area": f"{scrap_sub_area}",
                         "scrap_code": f"{scrap_code}",
                         "scrap_customer": f"{scrap_customer}",
                         "scrap_header": f"{scrap_header}",
                         "scrap_cost_center": f"{scrap_cost_center}",
                         }))

                except Exception as e:
                    error_window(e, traceback)
        for y in range(int(sap_instances)):
            thread_work = threading.Thread(target=do_work, name=str(y), args=(work,))
            thread_work.daemon = True
            all_threads.append(thread_work)
            thread_work.start()
        work.join()
        top.destroy()
        new_window()
    except Exception as e:
        error_window(e, traceback)


def mb1a_process():
    global files_
    t1 = threading.Thread(target=process_sap, daemon=True)
    files_ = fd.askopenfilenames(filetypes=[("Csv files", "*.csv")])
    if len(files_) == 0:
        return None
    t1.start()


def terminate(root_, top):
    check_v_1.set(False)
    check_v_2.set(False)
    check_v_3.set(False)
    top.update()
    top.destroy()
    root_.deiconify()
    os.system(f'taskkill /im saplogon.exe /f')
    # SAP_Functions.terminate()
    root.quit()


def refresh():
    global file
    global environment
    global storage_location
    global sap_instances

    file = pandas.read_excel(project_file, sheet_name="CONFIG", dtype=str)
    environment = str(file["Environment"][0])
    storage_location = str(file["Storage_Location"][0]).replace(".0", "")
    if int(float(file["SAP_Instances"][0])) >= int(sap_instances):
        sap_instances = int(float(file["SAP_Instances"][0]))
    root.title(f'MB1A_MASS:     Env: {environment} | St.Loc: {storage_location} | SAP.Instances: {sap_instances}')


def help_file():
    try:
        if getattr(sys, 'frozen', False):
            test = re.sub(r"(.*\\).*", "\\1", sys.executable)
            help_f = f'{test}\\help\\{project_name}_HELP.docx'

        else:
            help_f = f'{os.path.dirname(__file__)}/help/{project_name}_HELP.docx'
        subprocess.Popen([help_f], shell=True)
    except Exception as e:
        error_window(e, traceback)


def about():
    try:
        top = tk.Toplevel(width=305, height=200)
        top.iconbitmap('./img/image.ico')
        top.title("About MB1A_MASS")

        center_tk_window.center_on_screen(top)
        top.resizable(False, False)
        top.lift()
        label = ttk.Label(top, image=img3)
        label.image = img3
        label.grid(row=0, column=0, padx=40, pady=0, sticky=tk.W)
        label_1 = ttk.Label(top, text="Version 1.0.0")
        label_1.grid(row=5, column=0)

        #####################
    except Exception as e:
        error_window(e, traceback)


def new_window():
    try:
        top_p = tk.Toplevel(width=305, height=300)
        top_p.iconbitmap('./img/image.ico')
        top_p.title("Proceso terminado")
        root.withdraw()
        center_tk_window.center_on_screen(top_p)
        top_p.resizable(False, False)
        top_p.lift()
        label = ttk.Label(top_p, image=img1)
        label.image = img1
        label.grid(row=0, column=0, padx=40, pady=0, sticky=tk.W)
        label_1 = ttk.Label(top_p, text="Procesado")
        label_1.grid(row=5, column=0)
        progress_bar = ttk.Progressbar(top_p, orient="horizontal", mode="determinate", maximum=100, value=100)
        progress_bar.grid(row=6, column=0)

        button = ttk.Button(top_p, text="Terminar", width=50, command=lambda: terminate(root, top_p))
        button.grid(row=8, column=0, sticky=tk.W, pady=10)
        #####################
    except Exception as e:
        error_window(e, traceback)


def error_window(e_, traceback_):
    root.withdraw()
    res = messagebox.showerror(f'Error: {inspect.stack()[1].function}', f'{e_} \n\n {traceback_.format_exc()}')
    if res:
        root.quit()


try:

    work = queue.Queue()
    results = queue.Queue()
    sap_queue = queue.Queue()
    excel_queue = queue.Queue()
    current_process = 0
    label_2 = None
    textbox = None
    all_threads = []
    files_ = []

    if getattr(sys, 'frozen', False):
        project_file = re.sub(r".exe$", "", re.sub(r".*\\", "", sys.executable))
        project_name = re.sub(r".exe$", "", re.sub(r".*\\", "", sys.executable))
        project_file = f'{project_file}.xlsx'
    else:
        project_file = f'{re.sub(r".py", "", os.path.basename(__file__))}.xlsx'
        project_name = f'{re.sub(r".py", "", os.path.basename(__file__))}'

    file = pandas.read_excel(project_file, sheet_name="CONFIG", dtype=str)
    environment = str(file["Environment"][0])
    storage_location = str(file["Storage_Location"][0]).replace(".0", "")
    scrap_order = str(file["Order"][0]).replace(".0", "")
    sap_instances = int(float(file["SAP_Instances"][0]))
    root = tk.Tk()
    root.title(f'MB1A_MASS:     Env: {environment} | St.Loc: {storage_location} | SAP.Instances: {sap_instances}')
    root.iconbitmap('./img/image.ico')
    tabControl = ttk.Notebook(root)

    menu_bar = tk.Menu(root)
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Refresh", command=refresh)
    menu_bar.add_cascade(label="File", menu=file_menu)

    help_menu = tk.Menu(menu_bar, tearoff=0)
    help_menu.add_command(label="MB1A_MASS Help", command=help_file)
    help_menu.add_command(label="About...", command=about)
    menu_bar.add_cascade(label="Help", menu=help_menu)

    root.config(menu=menu_bar)

    tab1 = ttk.Frame(tabControl)

    tab1_image = tk.PhotoImage(file=r"./img/trash.png")

    tabControl.add(tab1, text='Movimientos de SCRAP Con Formato', image=tab1_image, compound=tk.LEFT)
    tabControl.grid(column=0, row=0, sticky=tk.E + tk.W + tk.N + tk.S)

    style = ttk.Style()
    style.configure("TLabelframe.Label", font=("TkDefaultFont", 10, "bold"), foreground='blue')
    # Notebook tab size
    current_theme = style.theme_use()
    style.theme_settings(current_theme, {"TNotebook.Tab": {"configure": {"padding": [20, 3]}}})

    # adding image (remember image should be PNG and not JPG)
    img1 = tk.PhotoImage(file=r"./img/Tristone_logo.png").subsample(2, 2)
    img3 = tk.PhotoImage(file=r"./img/Tristone.png").subsample(2, 2)

    # setting image with the help of label
    ttk.Label(tab1, image=img1).grid(row=0, column=2, columnspan=5, rowspan=2, padx=5, pady=5)

    # this will create CheckBoxes
    check_v_1 = tk.BooleanVar()
    check_v_2 = tk.BooleanVar()
    check_v_3 = tk.BooleanVar()
    check_v_1.set(False)
    check_v_2.set(False)
    check_v_3.set(False)
    #
    lfw = ttk.LabelFrame(tab1, text="Instrucciones")
    lw1 = ttk.Label(lfw, text="Para poder continuar cerciorarse de tener los siguientes puntos                                                ")
    check_1 = ttk.Checkbutton(lfw, text="1.-Cerrar todas las ventanas de SAP abiertas", var=check_v_1, command=check_status, )
    check_2 = ttk.Checkbutton(lfw, text="2.-Verificar que los formatos a procesar esten correctos", var=check_v_2, command=check_status)
    check_3 = ttk.Checkbutton(lfw, text="3.-Una vez verificados, guardar y cerrar los formatos.", var=check_v_3, command=check_status)

    # grid method to arrange labels in respective
    # rows and columns as specified
    lfw.grid(row=6, column=2, columnspan=5, sticky=tk.W, pady=4, padx=10)
    lw1.grid(row=8, column=2, columnspan=5, sticky=tk.W, pady=4)
    check_1.grid(row=12, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_2.grid(row=13, column=3, columnspan=5, sticky=tk.W, pady=0)
    check_3.grid(row=14, column=3, columnspan=5, sticky=tk.W, pady=0)

    # button widget
    b1 = ttk.Button(tab1, text="Seleccionar Archivo(s)", width=50, command=lambda: mb1a_process())

    # arranging button widgets
    b1.grid(row=18, column=3, columnspan=5, sticky=tk.W, pady=10)

    # infinite loop which can be terminated
    # by keyboard or mouse interrupt
    center_tk_window.center_on_screen(root)
    root.resizable(False, False)
    root.mainloop()
except Exception as E:
    error_window(E, traceback)
