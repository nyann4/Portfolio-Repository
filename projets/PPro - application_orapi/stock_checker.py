from tksheet import Sheet
import pandas as pd
from tkinter import *
from math import floor
from google.oauth2 import service_account
import gspread
import os
import time
import datetime as dt
from datetime import datetime
import random
import warnings
import numpy as np
import webbrowser


warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)

'''Functions'''
def switch(indicator_label, page): #use to switch between each tab when we click on another one
    global data
    df_complete, original_length, last_file = get_table(df_volume, list_non_empty)
    data= df_complete.values.tolist()
    data.insert(0, df_complete.columns)
    for tab in tab_frame.winfo_children() :
        if isinstance(tab, Label):
            tab['bg'] = 'SystemButtonFace'
    indicator_label['bg'] = '#355b69'
    for frame in main_frame.winfo_children(): # destroy the previous window before displaying the new one
        if isinstance(frame, Frame) or isinstance(frame, Button):
            frame.destroy()
            window.update()
    page()

def display_mr_info_adress(choosen_adress):
    for frame in display_mr_info_frame.winfo_children(): # destroy the previous window before displaying the new one
        frame.destroy()
        window.update()
    denomination = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Dénomination"].values[0]
    ref_fourni = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Référence fournisseur"].values[0]
    quantity = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Quantité"].values[0]
    date = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Date"].values[0]
    type = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Type"].values[0]
    label_infotitle_del_mr = Label(display_mr_info_frame, text=f"Informations liées à l'enregistrement manuel :  {choosen_adress}", font=("Helvetica",14), bg="#dd7c41")
    label_infotitle_del_mr.grid(column=0, row=0, sticky="nsew")
    label_info1_del_mr = Label(display_mr_info_frame, text=f"{denomination}   |    ref: {ref_fourni}", font=("Helvetica",14), bg="#dd7c41")
    label_info1_del_mr.grid(column=0, row=1, sticky="nsew")

    label_info2_del_mr = Label(display_mr_info_frame, text=f"type :{type}   |   quantité : {quantity}   |   date : {date}", font=("Helvetica",14), bg="#dd7c41")
    label_info2_del_mr.grid(column=0, row=2, sticky="nsew")
    if choosen_adress.startswith("1"): 
        for widgets in delete_mr_button_frame.winfo_children():# destroy the previous window before displaying the new one
            if isinstance(widgets, Label) or isinstance(widgets, Button):
                widgets.destroy()
                window.update()
        validation_delete_mr_label = Label(delete_mr_button_frame,text="Voulez vous vraiment supprimer cet enregistrement ?", font=("Helvetica",16), bg="#bcc5ca")
        validation_delete_mr_label.grid(column=0, row=0, sticky="ew", columnspan=4)
        validation_delete_mr_button = Button(delete_mr_button_frame, text="Suppression", font=("Helvetica",22), bg="#e13213", command= lambda :remove_selected_mr())
        validation_delete_mr_button.grid(column=0, row=1, rowspan=1, sticky ="nsew", columnspan=4)

def manual_scale_display():
    print(manual_generation_slider.get())

def remove_selected_mr():
    global choosen_adressv2
    if delete_variable == "Adresse" :
        choosen_adressv2 = choosen_adress
    remove_list = [["", "", "", "Vide", ""]]
    idx = df_manual.loc[df_manual["Emplacement"]== choosen_adressv2].index.tolist()[0]+2
    manual_sheet.batch_clear([f"B{idx}:F{idx}"])
    manual_sheet.update(remove_list, f'B{idx}:F{idx}')
    for widgets in title_frame_mr.winfo_children(): # destroy the previous window before displaying the new one
            widgets.destroy()
            window.update()
    manual_record_window()

def display_mr_type_adresse(): #display adress or mr for the selected type
    global drop_type_delete_choice_frame, om, choosen_type
    choosen_type = clicked_delete_type.get()
    if choosen_type == "Vide":
        try :
            for widgets in drop_type_delete_choice_frame.winfo_children(): # destroy the previous window before displaying the new one
                if isinstance(widgets, OptionMenu) :
                    widgets.destroy()
                    window.update()
        except :
            pass
        drop_type_delete.config(bg="RED", activebackground="#958885")
    if choosen_type != "Vide":
        drop_type_delete.config(bg="#dfc0b9", activebackground="#958885")
        for widgets in display_mr_info_frame.winfo_children(): # destroy the previous window before displaying the new one
            widgets.destroy()
            window.update()
        mr_adress_line = df_manual.loc[df_manual["Type"] == choosen_type, ["Emplacement", "Allée"]]
        drop_type_delete_choice_frame = Frame(search_delete_adress)
        drop_type_delete_choice_frame.config(background='#baecff')
        for w in range(0,4):
            drop_type_delete_choice_frame.grid_columnconfigure(minsize=90, index = w)
        for w in range(0,4):
            drop_type_delete_choice_frame.grid_rowconfigure(minsize=30, index = w)
        drop_type_delete_choice_frame.grid(column=0, row=2, sticky="ewsn", columnspan=8)
        om = {}
        i = 0
        while i < len(mr_adress_line["Allée"].unique()) :
            om[f"clicked_delete_type_adress_{i}"] = StringVar()
            i+=1
        i = 0
        for line in mr_adress_line["Allée"].unique():
            adress_list_by_line_type = mr_adress_line.loc[mr_adress_line["Allée"] == line, "Emplacement"].values
            o_m = OptionMenu(drop_type_delete_choice_frame, om[f"clicked_delete_type_adress_{i}"],
                        *adress_list_by_line_type, command =lambda value:check_all_line_mr_delete(value))
            o_m.grid(row=0+floor(i/4),column=i-4*floor(i/4), sticky="ew")
            o_m["menu"].config(bg="#dfc0b9", activebackground="#958885")
            o_m.config(bg="#dfc0b9", fg="BLACK", activebackground="#958885", activeforeground="BLACK")
            o_m.config(font=("Helvetica",13))
            o_m_name = window.nametowidget(o_m.menuname)
            o_m_name.config(font=("Helvetica",13))
            om[f"clicked_delete_type_adress_{i}"].set(line)
            i+=1
        # drop_type_delete_choice_name = window.nametowidget(drop_type_delete_choice.menuname)
        # drop_type_delete_choice_name.config(font=("Helvetica",14))
        # drop_type_delete_choice.grid(column=0, row=0, sticky="ewsn", rowspan=2, columnspan=4)

def check_all_line_mr_delete(choosen_adress):
    global choosen_adressv2
    choosen_adressv2 = choosen_adress
    for frame in display_mr_info_frame.winfo_children(): # destroy the previous window before displaying the new one
        frame.destroy()
        window.update()
    denomination = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Dénomination"].values[0]
    ref_fourni = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Référence fournisseur"].values[0]
    quantity = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Quantité"].values[0]
    date = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Date"].values[0]
    type = df_manual.loc[df_manual["Emplacement"] == choosen_adress, "Type"].values[0]
    label_infotitle_del_mr = Label(display_mr_info_frame, text=f"Informations liées à l'enregistrement manuel :  {choosen_adress}", font=("Helvetica",14), bg="#dd7c41")
    label_infotitle_del_mr.grid(column=0, row=0, sticky="nsew")
    label_info1_del_mr = Label(display_mr_info_frame, text=f"Dénomination : {denomination}", font=("Helvetica",14), bg="#dd7c41")
    label_info1_del_mr.grid(column=0, row=1, sticky="nsew")
    label_info2_del_mr = Label(display_mr_info_frame, text=f"ref: {ref_fourni} | type :{type}   ", font=("Helvetica",14), bg="#dd7c41")
    label_info2_del_mr.grid(column=0, row=2, sticky="nsew")

    label_info3_del_mr = Label(display_mr_info_frame, text=f"quantité : {quantity}   |   date : {date}", font=("Helvetica",14), bg="#dd7c41")
    label_info3_del_mr.grid(column=0, row=3, sticky="nsew")
    if choosen_adress.startswith("1"):
        for widgets in delete_mr_button_frame.winfo_children():# destroy the previous window before displaying the new one
            if isinstance(widgets, Label) or isinstance(widgets, OptionMenu) or isinstance(widgets, Button):
                widgets.destroy()
                window.update()
        validation_delete_mr_label = Label(delete_mr_button_frame, text="Voulez vous vraiment supprimer cet enregistrement ?", font=("Helvetica",16), bg="#bcc5ca")
        validation_delete_mr_label.grid(column=0, row=0, sticky="ew", columnspan=4)
        validation_delete_mr_button = Button(delete_mr_button_frame, text="Suppression", font=("Helvetica",22), bg="#e13213", command= lambda :remove_selected_mr())
        validation_delete_mr_button.grid(column=0, row=1, sticky ="nsew", columnspan=4)


#function to update the searchbar when we are typing
def update_searchbar_mr(base_variable):
    my_list_mr_delete.delete(0, END)

    for item in base_variable  : 
        my_list_mr_delete.insert(END, item)

# function to insert the item that we click on to the entry box
def fillout_searchbar_mr(event):
    try : 
        df = base_variable
    except :
        # manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
        # df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
        df = df_manual.loc[df_manual["Type"]!= "Vide","Emplacement"].tolist()

    my_entry_adress_mr_delete.delete(0, END)
    if my_list_mr_delete.curselection() : #if there is a current selection, display the first adress pick
        my_entry_adress_mr_delete.insert(0, df[my_list_mr_delete.curselection()[0]])
        global choosen_adress
        choosen_adress = df[my_list_mr_delete.curselection()[0]]
        display_mr_info_adress(choosen_adress)

#function to update the displayed list depending on what we type and if we can find it inside the elements of the list
def check_searchbar_mr(event):
    typed = my_entry_adress_mr_delete.get()
    new_data = []
    global base_variable, df_manual
    for i in df_manual.loc[df_manual["Type"]!= "Vide","Emplacement"].tolist():
        new_data.append(i) # retrieve only the adress from the existant list used in the google sheet display
    if typed == '':
        base_variable = df_manual.loc[df_manual["Type"]!= "Vide","Emplacement"].tolist()
    else :
        base_variable  = []
        
        for item in new_data :
            if item.lower().startswith(typed.lower()) :
                base_variable.append(item)
    update_searchbar_mr(base_variable) # after adding the corresponding item from our typed, we can use update data again

#function to update the searchbar when we are typing
def update_searchbar_allad(base_variable_v1):
    my_list.delete(0, END)

    for item in base_variable_v1  : 
        my_list.insert(END, item)

# function to insert the item that we click on to the entry box
def fillout_searchbar_allad(event):
    try :
        df = base_variable_v1
    except :
        df = df_emp["Emplacement"].tolist()
    my_entry_adress.delete(0, END)
    if my_list.curselection() :
        my_entry_adress.insert(0, df[my_list.curselection()[0]])
        global choosen_adress_allad
        choosen_adress_allad = df[my_list.curselection()[0]]
        info1_label.config(text=choosen_adress_allad, bg="#42c564", borderwidth=1, relief=SOLID)
        dict_add_mr_info["Adresse"] = True
    check_mr_info()

#function to update the displayed list depending on what we type and if we can find it inside the elements of the list
def check_searchbar_allad(event):
    typed = my_entry_adress.get()
    global base_variable_v1
    new_data = []
    for i in df_emp["Emplacement"].tolist():
        new_data.append(i) # retrieve only the adress from the existant list used in the google sheet display
    if typed == '':
        base_variable_v1 = df_emp["Emplacement"]
    else :
        base_variable_v1  = []
        for item in new_data :
            if item.lower().startswith(typed.lower()) :
                base_variable_v1.append(item)
    update_searchbar_allad(base_variable_v1) # after adding the corresponding item from our typed, we can use update data again

def set_delete_variable():
    global delete_variable, my_list_mr_delete, my_entry_adress_mr_delete, clicked_delete_type, drop_type_delete
    delete_variable = clicked_delete.get()
    for widgets in search_delete_adress.winfo_children():# destroy the previous window before displaying the new one
        if isinstance(widgets, Listbox) or isinstance(widgets, OptionMenu) or isinstance(widgets, Button):
            widgets.destroy()
            window.update()
    try :
        for widgets in drop_type_delete_choice_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(widgets, OptionMenu) :
                widgets.destroy()
                window.update()
    except : 
        pass
    # need to set a search bar depending on the variable delete_variable
    global base_variable, df_manual
    manual_sheet = stock_checker.worksheet("Enregistrement manuel")
    manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
    df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
    if delete_variable == "Adresse" :
        label_infotitle_del_mr = Label(display_mr_info_frame, text="Informations liées à l'enregistrement manuel : ", font=("Helvetica",14), bg="#dd7c41")
        label_infotitle_del_mr.grid(column=0, row=0, sticky="nsew")
        my_entry_adress_mr_delete = Entry(search_delete_adress, font=("Helvetica",16), width=20)
        my_entry_adress_mr_delete.grid(column=0, row=1, columnspan=5)
        search_deletead_label.config(text="Entrez une adresse")
        my_list_mr_delete = Listbox(search_delete_adress, width=20, font=("Helvetica",16), height=5)
        my_list_mr_delete.grid(column=0, row=2, columnspan=5)
        update_searchbar_mr(df_manual.loc[df_manual["Type"] != "Vide", "Emplacement"])
        my_list_mr_delete.bind("<<ListboxSelect>>", fillout_searchbar_mr)
        my_entry_adress_mr_delete.bind("<KeyRelease>", check_searchbar_mr)
    elif delete_variable == "Type":
        try :
            del base_variable
        except :
            pass

        type_search_delete = df_manual["Type"].unique()
        clicked_delete_type = StringVar()
        clicked_delete_type.set(type_search_delete[0])
        search_deletead_label.config(text="Choisissez un type")
        drop_type_delete = OptionMenu(search_delete_adress, clicked_delete_type, *df_manual["Type"].unique())
        drop_type_delete["menu"].config(bg="#dfc0b9", activebackground="#958885")
        drop_type_delete.config(bg="#dfc0b9", fg="BLACK", activebackground="#958885", activeforeground="BLACK")
        drop_type_delete.config(font=("Helvetica",14))
        drop_type_delete_name = window.nametowidget(drop_type_delete.menuname)
        drop_type_delete_name.config(font=("Helvetica",14))
        drop_type_delete.grid(column=0, row=1, sticky="ewsn", rowspan=1, columnspan=4)
        select_type_delete_button = Button(search_delete_adress, image=image_set, command= lambda :display_mr_type_adresse())
        select_type_delete_button.grid(column=4, row=1, sticky ="nsew")

# function to set the type when we are adding a new manual record
def set_type():
    global manual_type, manual_denom, manual_reference, manual_quantity
    actual_type = clicked.get()
    manual_type = my_entry_type.get()
    if manual_type != "":
        actual_type = manual_type
    if actual_type != "Vide": #avoid the empty type
        info2_label.config(text=actual_type, bg="#42c564", borderwidth=1, relief=SOLID)
        dict_add_mr_info["Type"] = True
        if actual_type in df_manual["Type"].unique() and actual_type != "Autres" and actual_type !="Artmef0": #force to add a denomination
            info3_label.config(text="/", bg="#42c564", borderwidth=1, relief=SOLID) #green
            info4_label.config(text="/", bg="#42c564", borderwidth=1, relief=SOLID)
            info5_label.config(text="/", bg="#42c564", borderwidth=1, relief=SOLID) #when we set an "Artmef0" or "other"
            manual_reference, manual_quantity, manual_denom, manual_type  = "/", "/", "/", actual_type
            dict_add_mr_info["Denomination"], dict_add_mr_info["Quantité"], dict_add_mr_info["Référence"] = True, True, True
        elif actual_type not in df_manual["Type"].unique() or actual_type == "Autres" or actual_type =="Artmef0":
            info3_label.config(text="Dénomination", font=("Helvetica",16), bg="#bcc5ca")
            info4_label.config(text="Quantité", font=("Helvetica",16), bg="#bcc5ca") #reset in gray the labels
            info5_label.config(text="Référence", font=("Helvetica",16), bg="#bcc5ca")
            manual_reference, manual_quantity, manual_denom, manual_type  = "", "", "", actual_type
            dict_add_mr_info["Denomination"], dict_add_mr_info["Quantité"], dict_add_mr_info["Référence"] = False, False, False

    else :
        info2_label.config(text=actual_type, bg="#d55b40", borderwidth=1, relief=SOLID)
        dict_add_mr_info["Type"] = False
    check_mr_info()

# function to set the denomination 
def set_denom():
    global manual_denom
    actual_denom = clicked_2.get()
    manual_denom = my_entry_denom.get()
    if manual_denom != "": # avoid no denomination
        actual_denom = manual_denom
    if actual_denom != " " and actual_denom != "":
        info3_label.config(text=actual_denom, bg="#42c564", borderwidth=1, relief=SOLID) #green
        dict_add_mr_info["Denomination"] = True
    else :
        info3_label.config(text="Dénomination", bg="#d55b40", borderwidth=1, relief=SOLID)
    check_mr_info()

# function to set the reference
def set_reference():
    global manual_reference
    actual_reference = clicked_4.get()
    manual_reference = my_entry_reference.get()
    if manual_reference != "": # avoid no reference
        actual_reference = manual_reference
    if actual_reference != " " and actual_reference != "":
        info5_label.config(text=actual_reference, bg="#42c564", borderwidth=1, relief=SOLID) #green
        dict_add_mr_info["Référence"] = True
    else :
        info5_label.config(text="Référence", bg="#d55b40", borderwidth=1, relief=SOLID)# red
    check_mr_info()

# check all mr info each time there is a modification
def check_mr_info() :
    checked = 0
    for key, value in dict_add_mr_info.items() :
        if dict_add_mr_info[key] == True :
            checked +=1
    if checked == 5:
        validation_mr_label = Label(validation_mr_frame,text="Voulez vous valider ?", font=("Helvetica",16), bg="#bcc5ca")
        validation_mr_label.grid(column=1, row=0, sticky="ew", columnspan=6)
        validation_mr_button = Button(validation_mr_frame, text="Ajouter", font=("Helvetica",24), bg="#179b2d", command= lambda :validate_mr())
        validation_mr_button.grid(column=1, row=1, rowspan=1, sticky ="nsew", columnspan=6)
    elif checked < 5 :
        for widgets in validation_mr_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(widgets, Frame) or isinstance(widgets, Button) or isinstance(widgets, Label):
                widgets.destroy()
                window.update()
    
def validate_mr():
    manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
    df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
    df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad, "Dénomination"] = manual_denom
    df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad, "Quantité"] = manual_quantity
    df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad, "Type"] = manual_type
    df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad, "Référence fournisseur"] = manual_reference
    df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad, "Date"] = dt.datetime.today().strftime('%d-%m-%Y')
    today = dt.datetime.today().strftime('%d-%m-%Y')
    idx = df_manual.loc[df_manual["Emplacement"]== choosen_adress_allad].index.tolist()[0]+2
    adding_mr_list = [[manual_denom, manual_reference, manual_quantity, manual_type, today]]
    manual_sheet.batch_clear([f"B{idx}:F{idx}"])
    manual_sheet.update(adding_mr_list, f'B{idx}:F{idx}')
    for widgets in title_frame_mr.winfo_children(): # destroy the previous window before displaying the new one
        widgets.destroy()
        window.update()
    manual_record_window()

# function to set the quantity
def set_quantity():
    global manual_quantity
    actual_quantity = clicked_3.get()
    manual_quantity = my_entry_quantity.get()
    if manual_quantity != "": # avoid no quantity
        actual_quantity = manual_quantity
    if actual_quantity != " " and actual_quantity != "":
        info4_label.config(text=actual_quantity, bg="#42c564", borderwidth=1, relief=SOLID) #green
        dict_add_mr_info["Quantité"] = True
    elif not actual_quantity.isdigit() :
        info4_label.config(text="!Nombres!", bg="#d55b40", borderwidth=1, relief=SOLID)# red
        dict_add_mr_info["Quantité"] = False
    else :
        info4_label.config(text="Quantité", bg="#d55b40", borderwidth=1, relief=SOLID)# red
        dict_add_mr_info["Quantité"] = False
    check_mr_info()

def delete_mr(): #use to destroy previous widgets to display the "add" ones
    for widgets in title_frame_mr.winfo_children(): # destroy the previous window before displaying the new one
        if isinstance(widgets, Frame) or isinstance(widgets, Button) or isinstance(widgets, Label):
            widgets.destroy()
            window.update()
    global clicked_delete, search_delete_adress, display_mr_info_frame, search_deletead_label, delete_mr_button_frame
    #create a frame to select a type or an adress
    search_delete_adress = Frame(title_frame_mr)
    search_delete_adress.config(background='#baecff')
    for w in range(0,5):
        search_delete_adress.grid_columnconfigure(minsize=70, index = w)
    for w in range(0,3):
        search_delete_adress.grid_rowconfigure(minsize=60, index = w)
    search_delete_adress.grid(column=1, row=1, sticky="ew", columnspan=5, rowspan=3)
    search_deletead_label = Label(search_delete_adress,text="Type de recherche", font=("Helvetica",16), bg="#dd7c41", borderwidth=1, relief=SOLID)
    search_deletead_label.grid(column=0, row=0, sticky="ewsn", columnspan=5)

    #create a frame to display information on the select type or adresse
    display_mr_info_frame = Frame(title_frame_mr)
    display_mr_info_frame.config(background='#dd7c41', borderwidth=1, relief=SOLID)
    display_mr_info_frame.grid_columnconfigure(minsize=300, index=0)
    for w in range(0,4):
        display_mr_info_frame.grid_rowconfigure(minsize=52, index = w)
    display_mr_info_frame.grid(column=2, row=4, sticky="nsew", columnspan=6, rowspan=4)
    label_infotitle_del_mr = Label(display_mr_info_frame, text="Informations liées à l'enregistrement manuel :", font=("Helvetica",14), bg="#dd7c41")
    label_infotitle_del_mr.grid(column=0, row=0, sticky="nsew")

    title_delete_label = Label(title_frame_mr, text="Suppression", font=("Helvetica",16), bg="#af5b5b")
    title_delete_label.grid(column=1, row=0, sticky="ew", columnspan=8)

    variable_delete = ["Adresse", "Type"]
    clicked_delete = StringVar()
    clicked_delete.set(variable_delete[0]) 

    select_variable_delete_button = Button(title_frame_mr, image=image_set, command= lambda :set_delete_variable())
    select_variable_delete_button.grid(column=8, row=1, rowspan=1, sticky ="nsew")
    drop_variable_delete = OptionMenu(title_frame_mr, clicked_delete, *variable_delete)
    drop_variable_delete.config(font=("Helvetica",14))
    drop_variable_delete_name = window.nametowidget(drop_variable_delete.menuname)
    drop_variable_delete_name.config(font=("Helvetica",14))
    drop_variable_delete.grid(column=6, row=1, sticky="ewsn", rowspan=1, columnspan=2)

    delete_mr_button_frame = Frame(title_frame_mr)
    delete_mr_button_frame.grid(column=2, row=8, sticky="nsew", columnspan=6, rowspan=2)
    delete_mr_button_frame.config(background='#baecff')

    for w in range(0,4):
        delete_mr_button_frame.grid_columnconfigure(minsize=45, index = w)
    for w in range(0,2):
        delete_mr_button_frame.grid_rowconfigure(minsize=52, index = w)

#Calculate the difference between 2 time on an apply method
def time_diff(x):
    y = (datetime.today().date() - pd.Timestamp(x).date()).days
    return y

#generate a list of product to check using the last file STOCK01 download
def random_inventory_auto():
    global path, df_log, df_inv_list, dict_today_product, ri_results_date, daily_check_frame, tamp, today_product, new_product, prod_number
    try :
        last_date_ri_check = df_log["Date"].tail(1).values[0]
    except IndexError:
        last_date_ri_check = tamp
    if last_date_ri_check != str(pd.Timestamp.today().date()) and tamp != str(pd.Timestamp.today().date()) and last_file_sto["time_diff"] <= 0.2: #If we already generate the RI today
        tamp = str(pd.Timestamp.today().date())
        ri_results_date = 0
        stock01 = pd.read_excel(path, skiprows=2)
        stock01 = stock01.loc[~stock01["Emplacement"].isna()]
        stock01 = stock01.sort_values("Article")
        #Will create a list and add every adress that contain more than one different reference to avoid doing mistake while modifying a stock
        double_emplacement = []
        exclusion_list = ["DATTENTE","ZARMOADV", "ZONPRECO", "ZQUAIREC","ZSHOWROO"]
        stock_double = stock01.groupby("Emplacement", as_index=False).agg(quantity=("Emplacement", "count"), reference=("Article", "nunique"))
        for i in stock_double.loc[stock_double["reference"]!=1, "Emplacement"] :
            if i not in exclusion_list :
                double_emplacement.append(i)
        
        #getting the number of palette for each product
        pic_number = pd.DataFrame({"Article" :stock01.loc[stock01["Type"]== "PIC", "Article"].value_counts().keys(),
                                "Picking":stock01.loc[stock01["Type"]== "PIC", "Article"].value_counts().values})
        sto_number_prod = pd.DataFrame({"Article" :stock01.loc[stock01["Type"]== "STO", "Article"].value_counts().keys(),
                                "Stock_lourd":stock01.loc[stock01["Type"]== "STO", "Article"].value_counts().values})

        #transform into 1 the multiple identical article by emplacement cause by different lot_id
        pic_number.loc[pic_number["Picking"]>1, "Number_p"] = 1

        #merge the pic and emp dataframe
        prod_number = pd.merge(pic_number, sto_number_prod, on="Article", how="outer")
        prod_number = prod_number.fillna(0)
        #get the real value as number of pic + stock for each
        prod_number["number"] = prod_number["Picking"] + prod_number["Stock_lourd"]

        prod_number = prod_number.sort_values("number", ascending=False)


        low1, low2, medium1, medium2, high1, max1, date = [], [], [], [], [], [], []

        new_product = prod_number.loc[~prod_number["Article"].isin(df_inv_list["Article"].unique()),"Article"].values

        shared_number = pd.merge(prod_number, df_inv_list, how="inner", on="Article")

        shared_number["weight"] = pd.to_numeric(shared_number["weight"])
        shared_number["categ"] = "0"
        shared_number.loc[shared_number["number"] <3, "categ"] = "low"
        shared_number.loc[(shared_number["number"] >2) & (shared_number["number"] <5),"categ"] = "medium"
        shared_number.loc[(shared_number["number"] >4) & (shared_number["number"] <11),"categ"] = "high"
        shared_number.loc[shared_number["number"] >10,"categ"] = "max"
        shared_number["cumsum"] = 0

        shared_number.loc[shared_number["categ"] == "low", "cumsum"] = shared_number.loc[shared_number["categ"] == "low", "weight"].cumsum()
        shared_number.loc[shared_number["categ"] == "medium", "cumsum"] = shared_number.loc[shared_number["categ"] == "medium", "weight"].cumsum()
        shared_number.loc[shared_number["categ"] == "high", "cumsum"] = shared_number.loc[shared_number["categ"] == "high", "weight"].cumsum()
        shared_number.loc[shared_number["categ"] == "max", "cumsum"] = shared_number.loc[shared_number["categ"] == "max", "weight"].cumsum()
        shared_number = shared_number.filter(items=["Article", "categ", "weight", "cumsum", "last_check", "diff"])

        #create the random variable to find the product to check
        today_product = []
        first_low = random.randint(1,shared_number.loc[shared_number["categ"]=="low", "cumsum"].max())
        second_low = random.randint(1,shared_number.loc[shared_number["categ"]=="low", "cumsum"].max())
        while not second_low != first_low :
            second_low = random.randint(1,sum(shared_number.loc[shared_number["categ"]=="low", "weight"]))

        first_med = random.randint(1,shared_number.loc[shared_number["categ"]=="medium", "cumsum"].max())
        second_med = random.randint(1,shared_number.loc[shared_number["categ"]=="medium", "cumsum"].max())
        while not second_med != first_low :
            second_med = random.randint(1,sum(shared_number.loc[shared_number["categ"]=="medium", "weight"]))

        first_high = random.randint(1,shared_number.loc[shared_number["categ"]=="high", "cumsum"].max())

        first_max = random.randint(1,shared_number.loc[shared_number["categ"]=="max", "cumsum"].max())

        #get the product to check and add them to a daily list
        low_1 = shared_number.loc[(shared_number["categ"]=="low")&(shared_number["cumsum"] >= first_low) & (shared_number["weight"]!=0)].head(1)
        low_2 = shared_number.loc[(shared_number["categ"]=="low")&(shared_number["cumsum"] >= second_low)& (shared_number["weight"]!=0)].head(1)
        medium_1 = shared_number.loc[(shared_number["categ"]=="medium")&(shared_number["cumsum"] >= first_med)& (shared_number["weight"]!=0)].head(1)
        medium_2 = shared_number.loc[(shared_number["categ"]=="medium")&(shared_number["cumsum"] >= second_med)& (shared_number["weight"]!=0)].head(1)
        high_1 = shared_number.loc[(shared_number["categ"]=="high")&(shared_number["cumsum"] >= first_high)& (shared_number["weight"]!=0)].head(1)
        max_1 = shared_number.loc[(shared_number["categ"]=="max")&(shared_number["cumsum"] >= first_max)& (shared_number["weight"]!=0)].head(1)
        today_product.append(low_1["Article"].values[0])
        today_product.append(low_2["Article"].values[0])
        today_product.append(medium_1["Article"].values[0])
        today_product.append(medium_2["Article"].values[0])
        today_product.append(high_1["Article"].values[0])
        today_product.append(max_1["Article"].values[0])

        display_stock(stock01, today_product, double_emplacement)
        '''The article that have to be modified with the following creation of new_inv_list need first to be pass into "unverified" filter'''#0002
        #add the new product to the daily checked one (to set them up to 0 weight)
        

        date.append(str(pd.Timestamp.today().date()))
        low1.append(today_product[0])
        low2.append(today_product[1])
        medium1.append(today_product[2])
        medium2.append(today_product[3])
        high1.append(today_product[4])
        max1.append(today_product[5])

        #need to modify this, the result have to come from the toplevel result page
        result = [0,0,0,0,0,0]
        global df_logv2
        df_logv2 = pd.DataFrame({"Date" :date, "Low1":low1, "Low2":low2,
                                "Medium1":medium1, "Medium2":medium2, "High1":high1,
                                "Max1":max1, "Low1_result":result[0], "Low2_result":result[1],
                                "Medium1_result":result[2], "Medium2_result":result[3], "High1_result":result[4],
                                "Max1_result":result[5]})

        #modify the google sheet file
        '''This 2 queries have to be part of the window where we will send the result of the stock checking'''
        '''I have to modify the first query to only update the reference that has been verified and not just blindly all of them'''


        for p in today_product :
            dict_today_product[f"{p}"] = int(prod_number.loc[prod_number["Article"] == p, "number"].values[0])
        destroy_main_page()
        stock_quality_page()
    elif last_date_ri_check == str(pd.Timestamp.today().date()) or tamp == str(pd.Timestamp.today().date()):
        for widgets in daily_check_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(widgets, Label) or isinstance(widgets, Button):
                widgets.destroy()
                window.update()
        daily_generation_button = Button(daily_check_frame, text="Déjà généré aujourd'hui !", font=("Arial", int(font_size*1.5)), fg='black', bg="#9A0A0A", command=lambda: random_inventory_auto())
        daily_generation_button.grid(column=0, row=0, sticky="nsew", columnspan=3, rowspan=2)
        daily_send_ri_button = Button(daily_check_frame, text="Envoyer", font=("Arial", int(font_size*1.5)), fg='white', bg="#4d6f7b", command= lambda: stock_check_details())
        daily_send_ri_button.grid(column=3, row=0, sticky="nsew", columnspan=1, rowspan=2)
    elif last_file_sto["time_diff"] > 0.2:
        for widgets in daily_check_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(widgets, Label) or isinstance(widgets, Button):
                widgets.destroy()
                window.update()
        daily_generation_button = Button(daily_check_frame, text="Fichier pas à jour !", font=("Arial", int(font_size*1.5)), fg='black', bg="#9A0A0A", command=lambda: random_inventory_auto())
        daily_generation_button.grid(column=0, row=0, sticky="nsew", columnspan=3, rowspan=2)
        daily_send_ri_button = Button(daily_check_frame, text="Envoyer", font=("Arial", int(font_size*1.5)), fg='white', bg="#4d6f7b", command= lambda: stock_check_details())
        daily_send_ri_button.grid(column=3, row=0, sticky="nsew", columnspan=1, rowspan=2)


def add_mr(): #use to destroy previous widgets to display the "add" ones
    for widgets in title_frame_mr.winfo_children(): # destroy the previous window before displaying the new one
        if isinstance(widgets, Frame) or isinstance(widgets, Button) or isinstance(widgets, Label):
            widgets.destroy()
            window.update()

    global my_entry_adress, my_list, info1_label, clicked, info2_label, my_entry_type, my_entry_denom, validation_mr_frame
    global clicked_2, info3_label, clicked_3, info4_label, clicked_4, info5_label, my_entry_quantity, my_entry_reference, dict_add_mr_info

    validation_mr_frame = Frame(title_frame_mr)
    validation_mr_frame.config(background='#baecff')
    validation_mr_frame.grid(column=1, row=10, sticky="ew", columnspan=8)
    for w in range(0,9):
        validation_mr_frame.grid_columnconfigure(minsize=80, index = w)
    for w in range(0,2):
        validation_mr_frame.grid_rowconfigure(minsize=80, index = w)
    '''Adress set up and research'''

    dict_add_mr_info = {"Adresse": False,"Type":False, "Denomination":False, "Quantité":False, "Référence":False,
                        "Adresse_value":-1000,"Denomination_value":-1000,"Quantité_value":-1000,"Référence_value":-1000,"Type_value":-1000}
    title_add_label = Label(title_frame_mr, text="Ajouter", font=("Helvetica",16), bg="#57a361")
    title_add_label.grid(column=1, row=0, sticky="ew", columnspan=4)
    info1_label = Label(title_frame_mr,text="Adresse", font=("Helvetica",16), bg="#bcc5ca")
    info1_label.grid(column=1, row=1, sticky="ew", columnspan=2)
    info2_label = Label(title_frame_mr,text="Type", font=("Helvetica",16), bg="#bcc5ca")
    info2_label.grid(column=3, row=1, sticky="ew", columnspan=2)
    info3_label = Label(title_frame_mr,text="Dénomination", font=("Helvetica",16), bg="#bcc5ca")
    info3_label.grid(column=1, row=5, sticky="ew", columnspan=4)
    info4_label = Label(title_frame_mr,text="Quantité", font=("Helvetica",16), bg="#bcc5ca")
    info4_label.grid(column=1, row=8, sticky="ew", columnspan=2)
    info5_label = Label(title_frame_mr,text="Référence", font=("Helvetica",16), bg="#bcc5ca")
    info5_label.grid(column=3, row=8, sticky="ew", columnspan=2)

    my_entry_adress = Entry(title_frame_mr, font=("Helvetica",16), width=20)
    my_entry_adress.grid(column=1, row=3)
    search_label = Label(title_frame_mr,text="Entrez une adresse", font=("Helvetica",16))
    search_label.grid(column=1, row=2, sticky="ewsn")

    my_list = Listbox(title_frame_mr, width=20, font=("Helvetica",16), height=5)
    my_list.grid(column=1, row=4)
    update_searchbar_allad(df_emp["Emplacement"])
    my_list.bind("<<ListboxSelect>>", fillout_searchbar_allad)
    my_entry_adress.bind("<KeyRelease>", check_searchbar_allad)

    '''Type set up and research'''
    global types
        #request the api to have an accurate list of types
    manual_sheet = stock_checker.worksheet("Enregistrement manuel")
    manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
    cardboard_type = ["Palette carton 8x5L", "Palette carton 4x5L", "Palette carton 2x5L", "Palette carton 1x5L"]
    #Keep the exact syntax of the cardboard palette name to put in the choosing type list when adding a mr
    df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
    types =df_manual["Type"].unique() #that's here that I set up the types list #0001d
    for i in cardboard_type :
        if i in types :
            pass
        else :
            types=  np.append(types, i)    
    clicked = StringVar()
    clicked.set(df_manual["Type"].unique().tolist()[0]) 

    my_entry_type = Entry(title_frame_mr, font=("Helvetica",16), width=16)
    my_entry_type.grid(column=3, row=3)
    select_type_button = Button(title_frame_mr, image=image_set, command= lambda :set_type())
    select_type_button.grid(column=4, row=2, rowspan=2, sticky ="nsew")
    drop_types = OptionMenu(title_frame_mr, clicked, *types)
    drop_types.config(font=("Helvetica",14))
    drop_types_name = window.nametowidget(drop_types.menuname)
    drop_types_name.config(font=("Helvetica",14))
    drop_types.grid(column=3, row=2, sticky="ewsn")

    '''Denomination set up'''
    clicked_2 = StringVar() 
    my_entry_denom = Entry(title_frame_mr, font=("Helvetica",16))
    my_entry_denom.grid(column=1, row=7, columnspan=4)
    select_denom_button = Button(title_frame_mr, image=image_set, command= lambda :set_denom())
    select_denom_button.grid(column=4, row=6, rowspan=2, sticky ="ns")

    '''quantity setup'''
    clicked_3 = StringVar() 
    my_entry_quantity = Entry(title_frame_mr, font=("Helvetica",16))
    my_entry_quantity.grid(column=1, row=9)
    select_quantity_button = Button(title_frame_mr, image=image_set_small, command= lambda :set_quantity())
    select_quantity_button.grid(column=2, row=9, rowspan=1, sticky ="ns")

    '''reference setup'''
    clicked_4 = StringVar()
    my_entry_reference = Entry(title_frame_mr, font=("Helvetica",16))
    my_entry_reference.grid(column=3, row=9)
    select_reference_button = Button(title_frame_mr, image=image_set_small, command= lambda :set_reference())
    select_reference_button.grid(column=4, row=9, rowspan=1, sticky ="ns")

def accessing_empty():
    list_path_file = []
    download_path = f"C:\\Users\\{os.getlogin()}\\Downloads"
    try :
        files = os.listdir(download_path) # files in the specified folder
        for file in files: # loop for each file
            path_file = f"{download_path}\\{file}"
            if file.endswith("SEA.csv") : # if we have a csv file, we open it to check if it is the good one
                check_df = pd.read_csv(path_file, nrows=2, sep= ";") #check first line of the files to only get STO files
                if all(col_2 in check_df.columns for col_2 in ["Emplacement", "Type emp", "Occupation"]) : #check the columns
                    #Need to add a check level -> occupation =1 -> only empty stock
                    if check_df["Occupation"].head(1).values[0] == 1 :
                        list_path_file.append(path_file)
        sorted_files = sorted(list_path_file, key=os.path.getmtime)
        now = dt.datetime.fromtimestamp(time.time())
        then = dt.datetime.fromtimestamp(os.path.getctime(sorted_files[-1]))
        tdelta = abs((now - then).total_seconds()/-3600)
        last_file = {"path": sorted_files[-1], "time":list(time.strptime((time.ctime(os.path.getctime(sorted_files[-1]))))), "time_diff": tdelta }
    except FileNotFoundError :
        print("there is a directory_path problem")
    return last_file

def accessing_full(): #Access the last type emp full file
    list_path_file = []
    download_path = f"C:\\Users\\{os.getlogin()}\\Downloads"
    files = os.listdir(download_path) # files in the specified folder
    for file in files: # loop for each file
        path_file = f"{download_path}\\{file}"
        if file.endswith("SEA.csv") : # if we have a csv file, we open it to check if it is the good one
            check_df = pd.read_csv(path_file, nrows=2, sep= ";") #check first line of the files to only get STO files
            if all(col_2 in check_df.columns for col_2 in ["Article", "Désignation 1"]) : #check the columns
                list_path_file.append(path_file)
            elif all(col_2 in check_df.columns for col_2 in ["Emplacement", "Type emp", "Occupation"]):
                print("Thats the wrong one")
                pass #need replace the "time" last file by a message "Mauvais fichier téléchargé"

    if list_path_file :
        sorted_files = sorted(list_path_file, key=os.path.getmtime)
        now = dt.datetime.fromtimestamp(time.time())
        then = dt.datetime.fromtimestamp(os.path.getctime(sorted_files[-1]))
        tdelta = abs((now - then).total_seconds()/-3600)
        last_file_full = {"path": sorted_files[-1], "time":list(time.strptime((time.ctime(os.path.getctime(sorted_files[-1]))))), "time_diff": tdelta
                          ,"problem":  "ok"}
    else :
        last_file_full = {"path": 0, "time": 0, "time_diff": 0, "problem": "Mauvais fichier téléchargé"}
    return last_file_full

def accessing_stock01():
    global last_file_sto, missing_file, dict_today_product
    list_path_file, missing_file = [], False
    download_path = f"C:\\Users\\{os.getlogin()}\\Downloads"
    files = os.listdir(download_path) # files in the specified folder
    for file in files: # loop for each file
        path_file = f"{download_path}\\{file}"
        if file.startswith("STOCK01") : # if we have a csv file, we open it to check if it is the good one
            list_path_file.append(path_file)
    sorted_files = sorted(list_path_file, key=os.path.getmtime)
    now = dt.datetime.fromtimestamp(time.time())
    if sorted_files : 
        then = dt.datetime.fromtimestamp(os.path.getctime(sorted_files[-1]))
        tdelta = abs((now - then).total_seconds()/-3600)
        last_file_sto = {"path": sorted_files[-1], "time":list(time.strptime((time.ctime(os.path.getctime(sorted_files[-1]))))), "time_diff": tdelta }
    else :
        print("The file is not download.")
        missing_file, last_file_sto = True, {"path":"", "time":[0,0,0,0,0,0,0,0,0], "time_diff":1000000}
    last_file_sto = change_time(last_file_sto)

def manual_access_designationdb():
    last_file_full = accessing_full()
    text_designation_label = last_file_full["problem"]
    if last_file_full["problem"] == "ok" :
        colorv3 = "green"
    else :
        colorv3 = "red"

    #Read the designation DB
    df_designation= pd.DataFrame(designation_sheet.get_all_values()[1:], columns=designation_sheet.get_all_values()[0])

    #Read the type emp full file
    df_type_emp_full = pd.read_csv(path, sep=";")
    df_type_emp_full = df_type_emp_full.filter(items=['Article', 'Désignation 1'])
    df_type_emp_full = df_type_emp_full.groupby("Article").agg(Designation=('Désignation 1', 'first'))
    df_type_emp_full.sort_values('Article', ascending=True, inplace=True)

    #Transform into list the previous article of the designationbd and actual article of the file
    designation_article = df_type_emp_full.index.values.tolist()
    previous_article = df_designation["Article"].values.tolist()
    #Compare the 2 list and find the differences
    new_article = list(set(designation_article) - set(previous_article))
    #Isolate the new product
    df_type_emp_full = df_type_emp_full.loc[df_type_emp_full.index.isin(new_article)]
    df_type_emp_full["Article"] = df_type_emp_full.index
    df_type_emp_full.reset_index(drop=True)
    #Reorder the columns to be able to concatenate the dataframes
    df_type_emp_full = df_type_emp_full.filter(items=["Article", "Designation"])

    df_designation_update = pd.concat([df_designation, df_type_emp_full])
    df_designation_update.sort_values("Article", ascending=True, inplace=True)


    designations = df_designation_update.values.tolist()
    if designations :
        designation_sheet.batch_clear(["A2:B12000"])
        designation_sheet.update(designations, f'A2:B{len(designations)+1}')
        print(df_type_emp_full.head())

    #Request the Designation DB
    '''Already request when starting the application -> designationdb = designation_sheet -> maybe not, I need to check'''
    designationdb_data= designation_sheet.get_all_values()
    df_designation = pd.DataFrame(designationdb_data[1:], columns=designationdb_data[0])
    print(df_designation.head())
    #Update the sheet

    #Update the related df

    #Update the displayed designation

    # df_type_emp_full = df_type_emp_full.filter(items=['Article', 'Désignation 1'])
    # df_type_emp_full = df_type_emp_full.groupby("Article").agg(Designation=('Désignation 1', 'first'))

    # df_type_emp_full = pd.merge(df_type_emp_full, df_designation, on="Article", how="outer")
    # df_type_emp_full.drop(columns=("Designation_y"), inplace=True)

    # designations = df_emp.values.tolist()
    # if designations :
    #     designation_sheet.batch_clear(["A2:B12000"])
    #     designation_sheet.update(designations, f'A2:B{len(designations)+1}')


def change_time(last_file): #change the time to display it in a better way
    for i, time in enumerate(last_file["time"][1:5]) :
        if len(str(time)) <2:
            last_file["time"][i+1] = f"0{str(last_file["time"][i+1])}"
    return last_file

def manual_update_page(): #Function to update the displayed table if we download a new file
    global last_file, df_complete, original_length, dict_info_mr, types
    manual_sheet = stock_checker.worksheet("Enregistrement manuel")
    manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
    df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
    types =df_manual["Type"].unique()

    cardboard_type = ["Palette carton 8x5L", "Palette carton 4x5L", "Palette carton 2x5L", "Palette carton 1x5L"]
    #Keep the exact syntax of the cardboard palette name to put in the choosing type list when adding a mr
    for i in cardboard_type :
        if i in types :
            pass
        else :
            types = np.append(types, i)
    dict_info_mr = general_info_mr(df_manual)
    list_non_empty = get_non_empty(df_manual)
    df_complete, original_length, last_file = get_table(df_volume, list_non_empty)
    destroy_main_page()
    apply_filter()
    # auto_saver()

def manual_update_ri(): #Function to update the displayed table if we download a new file
    global last_file_sto, df_log, df_inv_list, dict_info_mr, missing_file
    inv_list = random_inv.worksheet("inv_list")
    log = random_inv.worksheet("log")
    df_log= pd.DataFrame(log.get_all_values()[1:], columns=log.get_all_values()[0])
    df_inv_list= pd.DataFrame(inv_list.get_all_values()[1:], columns=inv_list.get_all_values()[0])
    #need to append the function that will get the stats of the last inventory here
    destroy_main_page()
    stock_quality_page()



def auto_update_page(): #Function to update the displayed table if we download a new file
    if actual_page == "empty_page" :
        global last_file, df_complete, original_length, dict_info_mr
        manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
        df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
        dict_info_mr = general_info_mr(df_manual)
        list_non_empty = get_non_empty(df_manual)
        df_complete, original_length, last_file = get_table(df_volume, list_non_empty)
        destroy_main_page()
        apply_filter()
    window.after(6000000, auto_update_page)

def get_table(df_volume, list_non_empty):
    last_file = accessing_empty()
    last_file = change_time(last_file)
    def my_function(x, h): #function to create the "allée" column using the "Emplacement" one
        return x[1:h]
    global df_emp
    df_emp = pd.read_csv(last_file["path"], sep= ";")
    #retrieve only the empty position to get a shorter dataframe then create an "Allée" column to apply filter on it
    df_emp = df_emp.loc[(df_emp["Type emp"] == "STO") & (df_emp["Occupation"] == 1)].filter(items=['Emplacement']) #keep the empty
    df_emp = df_emp.loc[~df_emp["Emplacement"].isin(list_non_empty)]
    df_emp["Allée"] = df_emp["Emplacement"].apply(my_function, args = [3])
    df_emp["pos"] = df_emp["Emplacement"].str[5:-1]
    df_emp["pos"] = pd.to_numeric(df_emp["pos"])
    df_emp["Côté"] = "0"
    df_emp.loc[df_emp["pos"]%2==1, "Côté"] = "I"
    df_emp.loc[df_emp["pos"]%2==0, "Côté"] = "P"
    df_emp.loc[(df_emp["Allée"] == "16")& (df_emp["Côté"] == "I"), "Allée"] = "16I"
    df_emp.loc[(df_emp["Allée"] == "16")& (df_emp["Côté"] == "P"), "Allée"] = "16P"
    df_emp = df_emp.filter(items=["Emplacement", "Allée"])

    #apply modification and filter to the volume dataframe
    df_volume = df_volume.filter(items=["Emplacement", "Hauteur", "Allée", "Position", "Côté"])
    df_volume = df_volume.astype({"Allée" : str})
    df_volume.loc[(df_volume["Allée"] == "16")& (df_volume["Côté"] == "I"), "Allée"] = "16I"
    df_volume.loc[(df_volume["Allée"] == "16")& (df_volume["Côté"] == "P"), "Allée"] = "16P"
    df_volume = df_volume.loc[df_volume["Hauteur"].isin(["1 étage", "2 étages", "3 étages", "4 étages", "5 étages", "6 étages", "max"])]
    df_volume = df_volume.filter(items=["Emplacement", "Hauteur", "Côté"])

    df_concat = pd.merge(df_emp, df_volume, on="Emplacement", how="inner")
    original_length = len(df_concat)
    return df_concat, original_length, last_file

def get_non_empty(df_manual):
    df_manualv2 = df_manual.loc[df_manual["Type"]!= "Vide"] #filter the non empty position
    df_manualv2= df_manualv2.filter(items=["Emplacement"]) # only keep the "Emplacement" column
    list_non_empty = df_manualv2.values # transform to list
    list_non_empty = [x for xs in list_non_empty for x in xs]
    return list_non_empty

def general_info_mr(df_manual):
    dict_info_mr = {}
    dict_info_mr["number_artemf0"] = len(df_manual.loc[df_manual["Type"]=="Artmef0"])
    dict_info_mr["number_mr"] = len(df_manual.loc[df_manual["Type"]!= "Vide"])
    dict_info_mr["prop_mr"] = round((len(df_manual.loc[df_manual["Type"]!= "Vide"])/len(df_manual))*100,2)
    dict_info_mr["carton_1x5L"] = len(df_manual.loc[df_manual["Type"]=="Palette carton 1x5L"])
    dict_info_mr["carton_2x5L"] = len(df_manual.loc[df_manual["Type"]=="Palette carton 2x5L"])
    dict_info_mr["carton_4x5L"] = len(df_manual.loc[df_manual["Type"]=="Palette carton 4x5L"])
    dict_info_mr["carton_8x5L"] = len(df_manual.loc[df_manual["Type"]=="Palette carton 8x5L"])
    return dict_info_mr

def table_form(df): #function to read and modify the volume file -> need to be simplified to not restart all step everytime
    data= df.values.tolist()
    data.insert(0, df.columns)
    return data

def filter_data(df):
    for key, values in filter_dict.items() :
        if filter_dict[key] :
            df = df.loc[df[f"{key}"].isin(filter_dict[key])]
    return df

def apply_filter(): #apply filter that we select by accessing at the "lane" updated value and refreshing the table frame
    for frame in main_frame.winfo_children():
        if isinstance(frame, Frame) == gsheet_frame:
            frame.destroy()
            window.update()
    global data, df_filtred
    df_filtred = filter_data(df_complete)
    data = table_form(df_filtred)
    empty_page()
    if len(df_filtred) < original_length :
        send_gsheet()

def destroy_main_page():
    for frame in main_frame.winfo_children():
        if isinstance(frame, Frame) == main_frame:
            frame.destroy()
            window.update()

def send_empty_gsheet(): #will send the filtered empty range to a google sheet file to be able to print it
    data_empty = df_filtred.values.tolist()
    if data_empty :
        empty_sheet.batch_clear(["A2:D2000"])
        empty_sheet.update(data_empty, f'A2:D{len(data_empty)+1}')
        
def display_stock(stock01, today_product, double_emplacement):
    stock01v2 = stock01.copy()
    #merge the actual stock01 with the designation dataframe
    data_print_inv = pd.merge(stock01v2, df_designation, on="Article", how="left") #the problem may be the way to merge the df
    data_print_inv = data_print_inv.loc[data_print_inv["Article"].isin(today_product)] #only keep the today product
    data_print_inv.loc[data_print_inv["Designation"].isna(), "Designation"] = "// Designation manquante //"

    print(data_print_inv.columns)
    data_print_inv["Quantité US"] = data_print_inv["Quantité US"].astype("str")
    data_print_inv["Quantité UC"] = data_print_inv["Quantité UC"].astype("str")
    data_print_inv["D"] = " "
    data_print_inv.loc[data_print_inv["Emplacement"].isin(double_emplacement), "D"] = "X"
    data_print_inv = data_print_inv.filter(items=["Article" ,"Emplacement", "D","Designation","Quantité UC", "Unité", "Quantité US"])
    data_print_inv = data_print_inv.sort_values("Emplacement")
    data_print_inv = data_print_inv.values.tolist()
    
    if data_print_inv :
        print_stock_sheet.batch_clear(["A2:G100"])
        print_stock_sheet.update(data_print_inv, f'A2:G{len(data_print_inv)+1}')

# def auto_saver():
#     save_df = save_sheet.get_all_values()
#     save_df = pd.DataFrame(save_df[1:], columns=save_df[0])
#     today = dt.datetime.today().strftime('%Y-%m-%d')
#     if not save_df["Date"].tail(1).values[0] == today :
#         save_sheet.update_cell(value=today, col=1, row=len(save_df)+2)
#         stock_checker_J1 = file.open("Stock Checker_sauvegarde_J-1")
#         gestion_stock_J1 = file.open("Gestion de stock_sauvegarde_J-1")

#         new_data = volume_sheet.get_all_values()
#         volume_sheet_J1 = stock_checker_J1.worksheet("Volume de stock")
#         volume_sheet_J1.batch_clear(["A1:G4000"])
#         volume_sheet_J1.update(new_data, f'A1:G{len(new_data)+1}')

#         new_data = manual_sheet.get_all_values()
#         manual_sheet_J1 = stock_checker_J1.worksheet("Enregistrement manuel")
#         manual_sheet_J1.batch_clear(["A1:H4000"])
#         manual_sheet_J1.update(new_data, f'A1:J{len(new_data)+1}')

#         new_data = regulation_sheet.get_all_values()
#         regulation_sheet_J1 = gestion_stock_J1.worksheet("Régulation de stock") #1500
#         regulation_sheet_J1.batch_clear(["A1:J1500"])
#         regulation_sheet_J1.update(new_data, f'A1:J{len(new_data)+1}')

#         new_data = f2023_sheet.get_all_values()
#         f2023_sheet_J1 = gestion_stock_J1.worksheet("2023") #1000
#         f2023_sheet_J1.batch_clear(["A1:K1000"])
#         f2023_sheet_J1.update(new_data, f'A1:K{len(new_data)+1}')
#         new_data = f2024_sheet.get_all_values()
#         f2024_sheet_J1 = gestion_stock_J1.worksheet("2024") #1000
#         f2024_sheet_J1.batch_clear(["A1:K1000"])
#         f2024_sheet_J1.update(new_data, f'A1:K{len(new_data)+1}')
#         new_data = f2025_sheet.get_all_values()
#         f2025_sheet_J1 = gestion_stock_J1.worksheet("2025") #1000
#         f2025_sheet_J1.batch_clear(["A1:K1000"])
#         f2025_sheet_J1.update(new_data, f'A1:K{len(new_data)+1}')
#         new_data = f2026_sheet.get_all_values()
#         f2026_sheet_J1 = gestion_stock_J1.worksheet("2026") #1000
#         f2026_sheet_J1.batch_clear(["A1:K1000"])
#         f2026_sheet_J1.update(new_data, f'A1:k{len(new_data)+1}')


def on_click(filter_dict, variable, values, bool_values): #add every checked lane to a main list to filter the table
    filter_dict[f"{variable}"] = [values[i] for i, chk in enumerate(bool_values) if chk.get()]

def on_click_ri(results_dict, variable, values, main_list): #add every checked product to a main list to save the logs from stock checking
    global big_check, small_check, ok_check, no_check
    if variable == "big" : 
        bool_values = big_check
    elif variable == "small":
        bool_values =small_check
    elif variable == "ok" :
        bool_values = ok_check
    else :
        bool_values = no_check
    results_dict[f"{variable}"] = [values[i] for i, chk in enumerate(bool_values) if chk.get()]
    main_list = []
    for key, val in results_dict.items():
        main_list +=val
    #Destroy the previous button/widget then replace it with updated values
    for frame in ri_validation_frame.winfo_children(): # destroy the previous window before displaying the new one
        if isinstance(frame, Label) or isinstance(frame, Button):
            frame.destroy()
            window.update()
    #Will display a different message depending on how many product are selected or if there is double, if no problem, display validation button
    if len(np.unique(main_list)) == len(main_list) and len(main_list) == 6:
        details_ri_validation = Button(ri_validation_frame, text="Valider", font=("Arial", int(font_size*1.2)), fg='white', bg="#08851D", command= lambda: send_ri_results())
        details_ri_validation.grid(column=0, row=0, sticky="nsew", columnspan=2, rowspan=1)
    elif len(np.unique(main_list)) == len(main_list) and len(main_list) != 6:
        details_ri_problem = Label(ri_validation_frame, text="Sélectionnez tous les articles", font=("Arial", int(font_size*1.2)), fg='black', bg="#DEF510", borderwidth=1, relief="solid")
        details_ri_problem.grid(column=0, row=0, sticky="nsew", columnspan=2, rowspan=1)
    else :
        details_ri_problem = Label(ri_validation_frame, text="Une catégorie par article !", font=("Arial", int(font_size*1.2)), fg='black', bg="#EE0520", borderwidth=1, relief="solid")
        details_ri_problem.grid(column=0, row=0, sticky="nsew", columnspan=2, rowspan=1)

def send_ri_results():
    global df_log, df_logv2, ri_results_date, new_inv_list
    results_ri = []
    ri_checked = []
    for key, val in results_dict.items():
        for v in results_dict[key] :
            if key == "big":
                ri_checked.append(v)
                results_ri.append("3")
            elif key == "small":
                ri_checked.append(v)
                results_ri.append("2")
            elif key == "ok":
                ri_checked.append(v)
                results_ri.append("1")
            elif key == "unverified":
                results_ri.append("0")
    #Will modified all the defaults values of the result only when the 6 different product has been checked


    to_zero_weight = [*ri_checked, *new_product]

    new_inv_list = pd.merge(prod_number, df_inv_list, on="Article", how="left", suffixes=('', '_y'))
    new_inv_list.drop(new_inv_list.filter(regex='_y$').columns, axis=1, inplace=True)
    new_inv_list.loc[new_inv_list["last_check"].isna(), "last_check"] = pd.Timestamp.today().date()
    new_inv_list.loc[new_inv_list["cumsum"].isna(), "cumsum"] = 0
    new_inv_list["weight"] = pd.to_numeric(new_inv_list["weight"])

    #need to set the columns value for all the new products
    new_inv_list["categ"] = "0"
    new_inv_list.loc[new_inv_list["number"] <3, "categ"] = "low"
    new_inv_list.loc[(new_inv_list["number"] >2) & (new_inv_list["number"] <6),"categ"] = "medium"
    new_inv_list.loc[(new_inv_list["number"] >5) & (new_inv_list["number"] <11),"categ"] = "high"
    new_inv_list.loc[new_inv_list["number"] >10,"categ"] = "max"

    #add rules depending on the categ to increase weight after the last_check
    new_inv_list['last_check'] = new_inv_list["last_check"].astype("datetime64[ns]")
    new_inv_list['diff'] = new_inv_list["last_check"].apply(time_diff)

    # Use number of day and category to define when I will unlock weight variable to be iterate again
    new_inv_list.loc[new_inv_list['weight']>0, "weight"] +=1 # Should resolve the problem of weight that doesnt increase with time
    new_inv_list.loc[((new_inv_list['categ']=="low") & (new_inv_list['diff'] >27) |
                    (new_inv_list['categ']=="medium") & (new_inv_list['diff'] >20) |
                    (new_inv_list['categ']=="high") & (new_inv_list['diff'] >12) |
                    (new_inv_list['categ']=="max") & (new_inv_list['diff'] >5))
                    & (new_inv_list['weight']==0), "weight"] +=1

    new_inv_list.loc[new_inv_list["Article"].isin(to_zero_weight), "weight"] = 0
    new_inv_list.loc[new_inv_list["Article"].isin(to_zero_weight), "last_check"] = pd.Timestamp.today().date()

    new_inv_list = new_inv_list.filter(items=["Article", "categ", "weight", "cumsum", "last_check", "diff"])
    new_inv_list["last_check"] = new_inv_list["last_check"].astype("str")

    df_logv2["Low1_result"], df_logv2["Low2_result"], df_logv2["Medium1_result"] = results_ri[0], results_ri[1], results_ri[2]
    df_logv2["Medium2_result"], df_logv2["High1_result"], df_logv2["Max1_result"] = results_ri[3], results_ri[4], results_ri[5]
    df_log = pd.concat([df_log, df_logv2])

    if ri_results_date == 0:
        ri_results_date = pd.Timestamp.today().date()
        
        data_empty = new_inv_list.values.tolist()
        inv_list_sheet.batch_clear(["A2:F8000"])
        inv_list_sheet.update(data_empty, f'A2:F{len(data_empty)+1}')
        #modify the google sheet log file
        data_log = df_log.values.tolist()
        log_sheet.batch_clear(["A2:M8000"])
        log_sheet.update(data_log, f'A2:M{len(data_log)+1}')
        for widgets in title_frame_ri.winfo_children(): # destroy the previous window before displaying the new one
            widgets.destroy()
            window.update()
        ri_page.destroy()
    else :
        for frame in ri_validation_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(frame, Label) or isinstance(frame, Button):
                frame.destroy()
                window.update()
        details_ri_problem = Label(ri_validation_frame, text="Résultats déjà envoyés", font=("Arial", int(font_size*1.2)), fg='black', bg="#EEDB05", borderwidth=1, relief="solid")
        details_ri_problem.grid(column=0, row=0, sticky="nsew", columnspan=2, rowspan=1)
        # ri_page.destroy()
    

def select_all(bool_values, values, variable): # add a full list of lane if necessary, can unselect all too
    base = False
    for c in bool_values:
        if c.get() == base :
            base = True
            break
    for c in bool_values:
        c.set(base)
    filter_dict[f"{variable}"] = values

def my_function(x, h): #function to create the "allée" column using the "Emplacement" one
    return x[1:h]

def send_gsheet():
    number_filter = Button(lane_frame, text="Envoyer sur google sheet", font=("Arial", int(font_size*1.5)), fg='#000000', command= lambda : send_empty_gsheet())
    number_filter.grid(row=3, column=0)

def manual_record_sub_window():
    mr_page = Toplevel()

    mr_page.title("Enregistrement manuel")
    mr_page.geometry("800x640")
    mr_page.minsize(800, 680)
    mr_page.maxsize(800, 680)
    global title_frame_mr
    
    title_frame_mr = Frame(mr_page)
    title_frame_mr.config(background='#baecff', borderwidth=1, relief=FLAT)
    title_frame_mr.grid(column=0, row=0)
    for w in range(0,10):
        title_frame_mr.grid_columnconfigure(minsize=80, index = w)

    for w in range(0,17):
        title_frame_mr.grid_rowconfigure(minsize=40, index = w)
    manual_record_window()

def manual_record_window():
    main_title_mr = Label(title_frame_mr, text="Enregistrement Manuel", font=("Arial", int(font_size*1.5)), fg='black', bg="#6565fd", borderwidth=1, relief="solid")
    main_title_mr.grid(column=2, row=1, sticky="nsew", columnspan=6)

    add_mr_button = Button(title_frame_mr, text="Ajouter", font=("Arial", int(font_size*1.5)), fg='black', bg="#57a361", command= lambda : add_mr())#2f8726
    add_mr_button.grid(column=1, row=3, sticky="nsew", columnspan=3, rowspan=3)

    search_mr_button = Button(title_frame_mr, text="Indisponible", font=("Arial", int(font_size*1.5)), fg='black', bg="#8c9196") #ffd91e
    search_mr_button.grid(column=6, row=3, sticky="nsew", columnspan=3, rowspan=3)

    modify_move_mr_button = Button(title_frame_mr, text="Indisponible", font=("Arial", int(font_size*1.5)), fg='black', bg="#8c9196") #2a6ea4
    modify_move_mr_button.grid(column=1, row=7, sticky="nsew", columnspan=3, rowspan=3)

    delete_mr_button = Button(title_frame_mr, text="Supprimer", font=("Arial", int(font_size*1.5)), fg='black', bg="#af5b5b", command= lambda : delete_mr())#de1717
    delete_mr_button.grid(column=6, row=7, sticky="nsew", columnspan=3, rowspan=3)

    '''main info manual records'''
    
    number_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["number_mr"]} enregistrement manuels", font=("Arial", int(font_size*1.1)), fg='black', bg="#baecff")
    number_mr_label.grid(column=1, row=11, sticky="nsew", columnspan=3, rowspan=1)
    prop_mr_label = Label(title_frame_mr, text=f"Soit environ {dict_info_mr["prop_mr"]}% de l'entrepôt", font=("Arial", int(font_size*1.1)), fg='black', bg="#baecff")
    prop_mr_label.grid(column=1, row=12, sticky="nsew", columnspan=3, rowspan=1)
    artmef0_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["number_artemf0"]} Artmef0 en Stock lourd", font=("Arial", int(font_size*1.1)), fg='black', bg="#baecff")
    artmef0_mr_label.grid(column=1, row=13, sticky="nsew", columnspan=3, rowspan=1)

    carton_1x5L_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["carton_1x5L"]} palettes de cartons 1x5L", font=("Arial", int(font_size*1.2)), fg='red', bg="#baecff")
    carton_1x5L_mr_label.grid(column=6, row=11, sticky="nsew", columnspan=3, rowspan=1)
    carton_2x5L_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["carton_2x5L"]} palettes de cartons 2x5L", font=("Arial", int(font_size*1.2)), fg='red', bg="#baecff")
    carton_2x5L_mr_label.grid(column=6, row=12, sticky="nsew", columnspan=3, rowspan=1)
    carton_4x5L_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["carton_4x5L"]} palettes de cartons 4x5L", font=("Arial", int(font_size*1.2)), fg='red', bg="#baecff")
    carton_4x5L_mr_label.grid(column=6, row=13, sticky="nsew", columnspan=3, rowspan=1)
    carton_8x5L_mr_label = Label(title_frame_mr, text=f"{dict_info_mr["carton_8x5L"]} palettes de cartons 8x5L", font=("Arial", int(font_size*1.2)), fg='red', bg="#baecff")
    carton_8x5L_mr_label.grid(column=6, row=14, sticky="nsew", columnspan=3, rowspan=1)

'''Functions to recreate each tab every time we change'''
def double_page(): #creation of the double emplacement frame that will be set by default
    actual_page = "double_page"
    double_frame = Frame(main_frame)
    double_frame.grid(row=0, column=0, sticky="ew")
    double_label_inf = Label(double_frame, text="Doubles emplacements", font=("Arial", int(font_size*1.5)), fg='black')
    double_label_inf.grid(sticky="ew")

def empty_page(): #creation of the empty emplacement frame that will be set by default
    global actual_page
    actual_page = "empty_page"
    global filter_dict
    filter_dict = {"Allée":[], "Hauteur":[]}
    lanes =["02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "13", "14", "15", "16P", "16I"]
    lanes_check = [BooleanVar() for i in lanes]

    sides =["Paire", "Impaire"]
    sides_check = [BooleanVar() for i in sides]

    heights =["1 étage", "2 étages", "3 étages", "4 étages", "5 étages", "6 étages", "max"]
    heights_check = [BooleanVar() for i in heights]

    '''first column widgets'''
    global lane_frame
    lane_frame = Frame(main_frame)
    lane_frame.grid(row=0, column=0, sticky="nsew")  
    lane_frame.grid_columnconfigure(minsize=320, index=0)

    lane_label = Label(lane_frame, text="Filtre par Allée", font=("Arial", int(font_size*1.8)), fg='#ffffff', bg='#619aaf', borderwidth=2, relief="solid")
    lane_label.grid(row=0, column=0, pady=5, padx=3, sticky="nsew")

    all_lane = Button(lane_frame, text="Tout sélectionner" , font=("Arial", int(font_size*1.5)), fg='#000000', bg="#c8cfd2",borderwidth=2, relief="raised",
                    command=lambda: select_all(lanes_check, lanes, "Allée"))
    all_lane.grid(row=1, column=0)

    checkboxlane_frame = Frame(lane_frame)
    checkboxlane_frame.grid(row=2, column=0)
    #creating a checkbox for each lane of the wharehouse
    for i, s in enumerate(lanes):
        Checkbutton(checkboxlane_frame, text=s, variable=lanes_check[i],
                     font=("Arial", int(font_size*1.5)), command=lambda : on_click(filter_dict, "Allée",lanes, lanes_check)).grid(row=1+floor(i/4),
                                                                                                              column=i-4*floor(i/4), sticky="ew")
    '''Second column widgets'''
    global gsheet_frame
    gsheet_frame = Frame(main_frame) #Frame containing the google sheet table and the filter button
    gsheet_frame.grid(row=0, column=1, sticky="nsew")

    side_label = Label(gsheet_frame, text="Filtre par côté", font=("Arial", int(font_size*1.8)), fg='#ffffff', bg='#619aaf', borderwidth=2, relief="solid")
    side_label.grid(row=0, column=0, pady=5, padx=3, sticky="nsew")

    side_frame = Frame(gsheet_frame)
    side_frame.grid(row=1, column=0)
    even_side = Checkbutton(side_frame, text="Paire", variable=sides_check[0], font=("Arial", int(font_size*1.5)), fg='#000000', command=lambda: on_click(filter_dict, "Côté",sides, sides_check))
    even_side.grid(row=0, column=0)
    all_side = Button(side_frame, text="Tout sélectionner", font=("Arial", int(font_size*1.5)), fg='#000000', bg="#c8cfd2",borderwidth=2, relief="raised",
                      command=lambda: select_all(sides_check, sides, "Côté"))
    all_side.grid(row=0, column=1)
    odd_side = Checkbutton(side_frame, text="Impaire", variable=sides_check[1], font=("Arial", int(font_size*1.5)), fg='#000000', command=lambda: on_click(filter_dict, "Côté", sides, sides_check))
    odd_side.grid(row=0, column=2)

    main_filter = Button(gsheet_frame, text="Filtrer", font=("Arial", int(font_size*1.6)), fg='#ffffff',bg="#355b69", command=lambda: apply_filter())
    main_filter.grid(row=2, column=0, sticky="nsew")
    main_filter.grid_columnconfigure(minsize=200, index=0)

    class table():
        def __init__(self):
            self.frame = Frame(gsheet_frame)
            self.frame.grid_columnconfigure(0, weight = 0)
            self.frame.grid_rowconfigure(0, weight = 0)
            self.sheet = Sheet(self.frame,
                            data = data, width=420)
            self.sheet.enable_bindings()
            self.frame.grid(row = 3, column = 0)
            self.sheet.grid(row = 0, column = 0)
    emp = table()

    manual_record_frame = Frame(main_frame) #Frame containing the google sheet table and the filter button
    manual_record_frame.grid(row=1, column=1, sticky="nsew")
    for w in range(0,3):
        manual_record_frame.grid_columnconfigure(w, weight = 1)
    manual_record_button = Button(manual_record_frame, text="Enregistrement manuel", font=("Arial", int(font_size*1.8)), fg='#ffffff',
     bg='#619aaf', borderwidth=2, relief="solid", command=manual_record_sub_window)
    manual_record_button.grid(row=0, column=1, sticky="nsew")


    '''third column widgets'''
    height_frame = Frame(main_frame)
    height_frame.grid(row=0, column=2, sticky="nsew")
    height_frame.grid_columnconfigure(minsize=320, index=0)

    height_label = Label(height_frame, text="Filtre par hauteur", font=("Arial", int(font_size*1.8)), fg='#ffffff', bg='#619aaf', borderwidth=2, relief="solid")
    height_label.grid(row=0, column=0, pady=5, padx=3, sticky="nsew")

    height_side_frame = Frame(height_frame)
    height_side_frame.grid(row=1, column=0)
    all_height = Button(height_side_frame, text="Tout sélectionner", font=("Arial", int(font_size*1.5)), fg='#000000', bg="#c8cfd2",borderwidth=2, relief="raised",
                        command=lambda: select_all(heights_check, heights, "Hauteur"))
    all_height.grid(row=0, column=0)


    update_button = Button(height_side_frame, image=image_update, font=("Arial", int(font_size*1.5)), fg='#000000', bg="#a1d1e3",borderwidth=2, relief="raised",
                        command=lambda: manual_update_page())
    update_button.grid(row=0, column=1, sticky="w", padx=20)


    checkboxheight_frame = Frame(height_frame)
    checkboxheight_frame.grid(row=2, column=0, columnspan=1)
    for i, s in enumerate(heights):
        Checkbutton(checkboxheight_frame, text=s, variable=heights_check[i],
                     font=("Arial", int(font_size*1.5)), command=lambda : on_click(filter_dict, "Hauteur",heights, heights_check)).grid(row=1+floor(i/2),
                                                                                                              column=i-2*floor(i/2), sticky="ew")
    
    empty_file_time = f"Date du fichier :{last_file["time"][2]}/{last_file["time"][1]}/{last_file["time"][0]}  {last_file["time"][3]}h{last_file["time"][4]}"
    if last_file["time_diff"] >=2 :
        color = "#FF0000"
    else :
        color = "#555556"
    file_time = Label(height_frame, text=empty_file_time, font=("Arial", int(font_size*1.2)), fg=color)
    file_time.grid(row=3, column=0, sticky="ne", pady=20)

def home_page(): #creation of the home frame that will be set by default
    actual_page = "home_page"
    for w in range(0,5):
        main_frame.grid_columnconfigure(w, weight = 1)
    home_page_frame = Frame(main_frame)
    home_page_frame.grid(row=0, column=2, sticky="ew")
    for w in range(0,7):
        home_page_frame.grid_columnconfigure(w, weight = 1)

    home_page_label_inf = Label(home_page_frame, text="Stock Checker", font=("Arial", int(font_size*1.5)), fg='black')
    home_page_label_inf.grid(row=0, column=3, sticky="ew")

    center_home_frame = Frame(main_frame)
    for w in range(0,4):
        center_home_frame.grid_columnconfigure(w, weight = 1)
    center_home_frame.grid(row=1, column=2,sticky="ew")

    center_illust_label = Label(center_home_frame, image=image_inventory, font=("Arial", int(font_size*1.5)), fg='black')
    center_illust_label.grid(row=0, column=1,sticky="ew")

def emp_picking_page(): #creation of the empty picking frame
    actual_page = "emp_picking_page"
    emp_picking_frame = Frame(main_frame)
    emp_picking_frame.grid(row=0, column=0, sticky="ew")
    emp_picking_label_inf = Label(emp_picking_frame, text="Picking disponible", font=("Arial", int(font_size*1.5)), fg='black')
    emp_picking_label_inf.grid(row=0, column=0, sticky="ew")

def stock_check_details():
    global title_frame_ri, actual_product_check, ri_validation_frame, results_dict, ri_page
    ri_page = Toplevel()
    ri_page.title("Enregistrement manuel")
    ri_page.geometry("960x720")
    ri_page.minsize(960, 720)
    ri_page.maxsize(960, 720)
    main_list = []
    actual_product_check = [k for k, v in dict_today_product.items()]
    palette_product_check = [v for k, v in dict_today_product.items()]
    title_frame_ri = Frame(ri_page)
    title_frame_ri.config(background='#baecff', borderwidth=1, relief=FLAT)
    #Setup the grid
    title_frame_ri.grid(column=0, row=0)
    for w in range(0, 8) :
        title_frame_ri.grid_columnconfigure(minsize=120, index=w)
    for w in range(0, 18) :
        title_frame_ri.grid_rowconfigure(minsize=40, index=w)

    main_title_detail_ri = Label(title_frame_ri, text="Résultats des vérifications", font=("Arial", int(font_size*1.4)), fg='black', bg="#6565fd", borderwidth=1, relief="solid")
    main_title_detail_ri.grid(column=2, row=1, sticky="nsew", columnspan=4, rowspan=1)

    ri_headers = ["Références","Palettes","+++ / - - -"," + / - ", "Stock juste", "Non-vérifié"]
    colors = ["#355b69", "#355b69", "#D42120", "#D4A120", "#74D420", "#9E9EA8"]
    for j, i in enumerate(ri_headers)  :
        ri_title_header = Label(title_frame_ri, text=i, font=("Arial", int(font_size*1.2)), fg='white', bg=colors[j], borderwidth=1, relief="solid")
        ri_title_header.grid(column=1+j, row=3, sticky="nsew", columnspan=1, rowspan=1)

    #List of product to replace either by "today_product" or by "manual_product"
    
    for j, i in enumerate(actual_product_check)  :
        references_detail_ri = Label(title_frame_ri, text=i, font=("Arial", int(font_size*1.2)), fg='white', bg="#407A94", borderwidth=1, relief="solid")
        references_detail_ri.grid(column=1, row=4+j, sticky="nsew", columnspan=1, rowspan=1)
    for j, i in enumerate(palette_product_check)  :
        palettes_detail_ri = Label(title_frame_ri, text=i, font=("Arial", int(font_size*1.2)), fg='white', bg="#407A94", borderwidth=1, relief="solid")
        palettes_detail_ri.grid(column=2, row=4+j, sticky="nsew", columnspan=1, rowspan=1)

    
    global big_check, small_check, ok_check, no_check
    big_check = [BooleanVar(value=False) for i in actual_product_check]
    small_check = [BooleanVar(value=False) for i in actual_product_check]
    ok_check = [BooleanVar(value=False) for i in actual_product_check]
    no_check = [BooleanVar(value=False) for i in actual_product_check]
    results_dict = {"big": [], "small":[], "ok" : [], "ok":[]}
    for j in range(0,len(actual_product_check)):
        Checkbutton(title_frame_ri, variable=big_check[j],
                    font=("Arial", int(font_size*1.2)), bg="#8FB1C4",
                    command=lambda : on_click_ri(results_dict, "big", actual_product_check, main_list)).grid(row=4+j, column=3, sticky="nsew")
    for j in range(0,len(actual_product_check)):
        Checkbutton(title_frame_ri, variable=small_check[j],
                    font=("Arial", int(font_size*1.2)), bg="#8FB1C4",
                    command=lambda : on_click_ri(results_dict, "small", actual_product_check, main_list)).grid(row=4+j, column=4, sticky="nsew")
    for j in range(0,len(actual_product_check)):
        Checkbutton(title_frame_ri, variable=ok_check[j],
                    font=("Arial", int(font_size*1.2)), bg="#8FB1C4",
                    command=lambda : on_click_ri(results_dict, "ok", actual_product_check, main_list)).grid(row=4+j, column=5, sticky="nsew")
    for j in range(0,len(actual_product_check)):
        Checkbutton(title_frame_ri, variable=no_check[j],
                    font=("Arial", int(font_size*1.2)), bg="#8FB1C4",
                    command=lambda : on_click_ri(results_dict, "unverified", actual_product_check, main_list)).grid(row=4+j, column=6, sticky="nsew")

    ri_validation_frame = Frame(title_frame_ri)
    ri_validation_frame.config(background="#baecff") #'#baecff'
    ri_validation_frame.grid(row=5+len(actual_product_check), column=3, sticky="nsew", columnspan=2)  
    ri_validation_frame.grid_columnconfigure(minsize=120, index=0)
    ri_validation_frame.grid_columnconfigure(minsize=120, index=1)
    ri_validation_frame.grid_rowconfigure(minsize=40, index=0)
    

    # Function Declaration
def callback():
    webbrowser.open_new(r"https://docs.google.com/spreadsheets/d/19SABMkaYvybF1ZtatyX1jgiEctUh9_vS3d8ufCm9DSE/edit?gid=37960476#gid=37960476")

def stock_quality_page(): #creation of the quality frame
    global last_file_sto, path, df_log, df_inv_list, dict_today_product, daily_check_frame
    quality_frame = Frame(main_frame)
    quality_frame.config(borderwidth=1, relief=FLAT)
    quality_frame.grid(column=0, row=0)
    x= 1
    for w in range(0,13):
        if w == 11 :
            s = 20
        else : 
            s = 82
        quality_frame.grid_columnconfigure(minsize=s, index = w)

    for w in range(0,22):
        quality_frame.grid_rowconfigure(minsize=25, index = w)

    daily_check_frame = Frame(quality_frame)
    daily_check_frame.config(borderwidth=1, relief=FLAT)
    daily_check_frame.grid(column=2, row=4+x, columnspan=4, rowspan=2)
    for w in range(0,4):
        daily_check_frame.grid_columnconfigure(minsize=82, index = w)
    for w in range(0,2):
        daily_check_frame.grid_rowconfigure(minsize=25, index = w)
    #Accessing the first time to the STO file to access the date of the last file
    accessing_stock01()

    #Add today product if they are already generate, or fill with
    product_check, number_stock = [], []
    if not dict_today_product :
        product_check, number_stock = [0,0,0,0,0,0], [0,0,0,0,0,0]
        color2 = "red"
        daily_generation_button = Button(daily_check_frame, text="Génération quotidienne", font=("Arial", int(font_size*1.5)), fg='white', bg="#355b69", command=lambda: random_inventory_auto())
        daily_generation_button.grid(column=0, row=0, sticky="nsew", columnspan=4, rowspan=2)
    else :
        for key, value in dict_today_product.items():
            product_check.append(key)
            number_stock.append(value)
        for widgets in daily_check_frame.winfo_children(): # destroy the previous window before displaying the new one
            if isinstance(widgets, Label) or isinstance(widgets, Button):
                widgets.destroy()
                window.update()
        daily_generation_button = Button(daily_check_frame, text="Génération quotidienne", font=("Arial", int(font_size*1.5)), fg='white', bg="#355b69", command=lambda: random_inventory_auto())
        daily_generation_button.grid(column=0, row=0, sticky="nsew", columnspan=3, rowspan=2)
        daily_send_ri_button = Button(daily_check_frame, text="Envoyer", font=("Arial", int(font_size*1.5)), fg='white', bg="#4d6f7b", command= lambda: stock_check_details())
        daily_send_ri_button.grid(column=3, row=0, sticky="nsew", columnspan=1, rowspan=2)
        color2 = "black"
    daily_product_info = f"{product_check[0]} : {number_stock[0]} | {product_check[1]} : {number_stock[1]} \n {product_check[2]} : {number_stock[2]} | {product_check[3]} : {number_stock[3]} \n {product_check[4]} : {number_stock[4]} | {product_check[5]} : {number_stock[5]}"
    #Retrieve the path value from the last file
    path = last_file_sto["path"]
    #Change color of the font depending on the timediff -> yellow if there is no file download
    if last_file_sto["time_diff"] >=0.2 :
        if last_file_sto["time_diff"] == 1000000:
            color = "#FFEA00"
        else :
            color = "#FF0000"
    else :
        color = "#555556"
    sto_time = f"Date :{last_file_sto["time"][2]}/{last_file_sto["time"][1]}/{last_file_sto["time"][0]}  {last_file_sto["time"][3]}h{last_file_sto["time"][4]}"
    
    global manual_generation_slider

    date_file_ri = Label(quality_frame, text=sto_time, font=("Arial", int(font_size*1.2)), fg=color)
    date_file_ri.grid(column=0*1, row=x, sticky="nsew", columnspan=3, rowspan=2)

    main_title_ri = Label(quality_frame, text="Génération d'inventaire aléatoire", font=("Arial", int(font_size*1.6)), fg='white', bg="#6565fd", borderwidth=2, relief="solid")
    main_title_ri.grid(column=4, row=x, sticky="nsew", columnspan=5, rowspan=2)

    update_stockfile_button = Button(quality_frame, image=image_update, font=("Arial", int(font_size*1.5)), fg='#000000', bg="#a1d1e3",borderwidth=2, relief="raised",
                                     command=lambda : manual_update_ri())
    update_stockfile_button.grid(column=10, row=x, sticky="nsew", columnspan=1, rowspan=2)
    #button to access the last full SEA file download and update the designation database
    designation_button = Button(quality_frame, image=image_arrow, font=("Arial", int(font_size*1.5)), fg='#000000', bg="#a1d1e3",borderwidth=2, relief="raised",
                                     command=lambda : manual_access_designationdb()) #Working here #0004
    designation_button.grid(column=12, row=x, sticky="nsew", columnspan=1, rowspan=2)
    #Label for the typ emp full last file
    last_file
    type_emp_full_file = Label(quality_frame, text="Ok", font=("Arial", int(font_size*1.3)), fg=color)
    type_emp_full_file.grid(column=12, row=x+2, sticky="nsew", columnspan=2, rowspan=2)

    daily_generation_label = Label(quality_frame, text=daily_product_info,  font=("Arial", int(font_size*1.3)), fg=color2)
    daily_generation_label.grid(column=2, row=7+x, sticky="nsew", columnspan=4, rowspan=5)

    manual_generation_button = Button(quality_frame, text="Génération manuelle", font=("Arial", int(font_size*1.5)), fg='white', bg="#355b69", command=manual_scale_display)
    manual_generation_button.grid(column=7, row=4+x, sticky="nsew", columnspan=3, rowspan=2)

    manual_send_ri_button = Button(quality_frame, text="Envoyer", font=("Arial", int(font_size*1.5)), fg='white', bg="#4d6f7b", command= lambda: stock_check_details())
    manual_send_ri_button.grid(column=10, row=4+x, sticky="nsew", columnspan=1, rowspan=2)

    manual_generation_slider  = Scale(quality_frame, from_=1, to=20, fg='black', orient=HORIZONTAL)
    manual_generation_slider.grid(column=7, row=6+x, sticky="nsew", columnspan=4, rowspan=1)

    check_ri_results_button = Button(quality_frame, text="Résultats d'inventaire aléatoire", font=("Arial", int(font_size*1.5)), fg='white', bg="#4d6f7b", command=callback)
    check_ri_results_button.grid(column=4, row=13, sticky="nsew", columnspan=5, rowspan=2)

# try :
'''File reading and adjusting''' # -> need to merge with the volume file and add the merged df to the table fram
scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

credentials = service_account.Credentials.from_service_account_file("key.json", scopes=scopes)

file = gspread.authorize(credentials)
stock_checker = file.open("Stock Checker")

volume_sheet = stock_checker.worksheet("Volume de stock")
empty_sheet = stock_checker.worksheet("Emplacement vide")
manual_sheet = stock_checker.worksheet("Enregistrement manuel")

gestion_stock = file.open("Gestion de stock")
random_inv = file.open("random_inventory")
save = file.open("log_sauvegarde")

regulation_sheet = gestion_stock.worksheet("Régulation de stock")


save_sheet = save.worksheet("sauvegarde")

#control stock google sheet access
inv_list_sheet = random_inv.worksheet("inv_list")
log_sheet = random_inv.worksheet("log")
print_stock_sheet = random_inv.worksheet("Liste à imprimer")
designation_sheet = random_inv.worksheet("DesignationDB")
df_log= pd.DataFrame(log_sheet.get_all_values()[1:], columns=log_sheet.get_all_values()[0])
df_inv_list= pd.DataFrame(inv_list_sheet.get_all_values()[1:], columns=inv_list_sheet.get_all_values()[0])
df_print_stock= pd.DataFrame(print_stock_sheet.get_all_values()[1:], columns=print_stock_sheet.get_all_values()[0])
df_designation = pd.DataFrame(designation_sheet.get_all_values()[1:], columns=designation_sheet.get_all_values()[0])


# total_sheet = gestion_stock.worksheet("Total")
# f2023_sheet = gestion_stock.worksheet("2023")
# f2024_sheet = gestion_stock.worksheet("2024")
# f2025_sheet = gestion_stock.worksheet("2025")
# f2026_sheet = gestion_stock.worksheet("2026")

manual_data = manual_sheet.get_all_values() #retrieve data from manual records sheet
df_manual = pd.DataFrame(manual_data[1:], columns=manual_data[0])
list_non_empty = get_non_empty(df_manual)
dict_info_mr = general_info_mr(df_manual)
data = volume_sheet.get_all_values() # Create data variable that will be use on the google sheet table displayed
df_volume = pd.DataFrame(data[1:], columns=data[0])
df_complete, original_length, last_file = get_table(df_volume, list_non_empty)

data= df_complete.values.tolist()
data.insert(0, df_complete.columns)

#declare this variable one time to fill with 0 when no daily check has been done
dict_today_product = {}


'''Variables'''
tamp = 0 
tab_width = 343
font_size = 11
padx = 10
numbe_of_tabs = 3
global actual_page
actual_page = "home_page"

'''Interface'''
window = Tk()


'''Creating image variable for different button'''
image_update = PhotoImage(file="app_image//update_icon.PNG")
image_inventory = PhotoImage(file="app_image//image_inventory.png")
image_set = PhotoImage(file="app_image//set_button.png")
image_set_small = PhotoImage(file="app_image//set_button_smaller.png")
image_arrow = PhotoImage(file="app_image//arrow_icon.png")

#personnalisation of the window
window.title("StockChecker")
window.config(background='#1b645a')
#size of the window
window.geometry(f"{numbe_of_tabs*tab_width+(numbe_of_tabs+1)*padx+10}x550")
window.minsize(numbe_of_tabs*tab_width+(numbe_of_tabs+1)*padx+10 ,550)
window.maxsize(numbe_of_tabs*tab_width+(numbe_of_tabs+1)*padx+10 ,550)

tab_frame = Frame(window)

# '''double emplacement tab'''
# double_label  =Button(tab_frame, text="Double emplacement",
#                         font = ("Arial", font_size), bg= '#355b69', fg='white',
#                         activeforeground='#002539', command=lambda: switch(indicator_label=double_indicator,
#                                                                             page=double_page))
# double_label.place(x=padx, y=0, width=tab_width)
# double_indicator = Label(tab_frame)
# double_indicator.place(x=25, y =35, width=170, height=2)

'''empty emplacement tab'''
empty_label  =Button(tab_frame, text="Emplacement vide",
                        font = ("Arial", font_size), bg= '#355b69', fg='white',
                        activeforeground='#002539', command=lambda: switch(indicator_label=empty_indicator,
                                                                            page=empty_page))
# empty_label.place(x=2*padx+tab_width, y=0, width=tab_width)
empty_label.place(x=padx+8, y=0, width=tab_width)
empty_indicator = Label(tab_frame)
empty_indicator.place(x=padx+20, y =35, width=313, height=2)

'''hometab'''
home_label  =Button(tab_frame, text="Page d'accueil",
                        font = ("Arial", font_size), bg= 'black', fg='white',
                        activeforeground='#002539', command=lambda: switch(indicator_label=home_indicator,
                                                                            page=home_page))
home_label.place(x=2*padx+1*tab_width+8, y=0, width=tab_width)
home_indicator = Label(tab_frame, bg='#355b69')
home_indicator.place(x=2*padx+1*tab_width+20, y =35, width=313, height=2)

'''Checking empty picking using stock03 file'''
# emp_picking_label  =Button(tab_frame, text="Picking disponibles",
#                         font = ("Arial", font_size), bg= '#355b69', fg='white',
#                         activeforeground='#002539', command=lambda: switch(indicator_label=emp_picking_indicator,
#                                                                             page=emp_picking_page))
# emp_picking_label.place(x=4*padx+3*tab_width, y=0, width=tab_width)
# emp_picking_indicator = Label(tab_frame)
# emp_picking_indicator.place(x=3*padx+3*tab_width+25, y =35, width=170, height=2)

'''quality tab'''
quality_label  =Button(tab_frame, text="Qualité des stocks",
                         font = ("Arial", font_size), bg= '#355b69', fg='white',
                         activeforeground='#002539', command=lambda: switch(indicator_label=quality_indicator,
                                                                      page=stock_quality_page))
quality_label.place(x=3*padx+2*tab_width+8, y=0, width=tab_width)
quality_indicator = Label(tab_frame)
quality_indicator.place(x=3*padx+2*tab_width+15, y =35, width=313, height=2)


tab_frame.grid(sticky="ew")

tab_frame.grid_propagate(False)
tab_frame.configure(width=numbe_of_tabs*362, height=38)


main_frame = Frame(window)
main_frame.grid(sticky="nsew")
home_page()
auto_update_page()
window.mainloop()