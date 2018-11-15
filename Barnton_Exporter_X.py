# -*- coding: utf-8 -*-

#import the pandas library
#import the keyboard simulator
#import time
import pandas as pd
from pynput.keyboard import Key, Controller
import time

while True:
    #Validate input
    while True:
        try:
            #Save the excel stocksheet to the variable 'file'.
            #The path is concatenated to the input in the backend to make it easier for the user to enter the name of the file.
            #"C:\Users\schikaura\Documents"
            file = input("Copy the name of the excel stocksheet into the space provided. Ensure the file is in Documents before you do this.\n")
            file = r"C:\Users\schikaura\Documents\\" + file + ".xlsx"
            data = pd.ExcelFile(file)
            break
        except FileNotFoundError as fnf_error:
            print(fnf_error)
            print("Please follow the instructions and type in the correct filename. If the file is not in the Documents folder, please copy it there first.")

    keyboard = Controller()

    while True:
        SBU_site_stocks = input("Please select the SBU/site for which you wish to export stocks for. \nEnter A for Harare Beer \nB for Bulawayo Beer \nC for Harare CSD \nD for Harare Maheu \nE for Bulawayo CSD \nF for Bulawayo PET")

        if SBU_site_stocks == "A":
            north_import = data.parse("North_import_sheet")

            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = north_import["Opening Stock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break

        elif SBU_site_stocks == "B":
            south_import = data.parse("South_import_sheet")
                
            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = south_import["OpeningStock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break

        elif SBU_site_stocks == "C":
            north_import = data.parse("N_Region_SBs_SKUs")
                
            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = north_import["OpeningStock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break
        
        elif SBU_site_stocks == "D":
            national_import = data.parse("Harare Maheu")
                
            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = national_import["OpeningStock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break

        elif SBU_site_stocks == "E":
            south_import = data.parse("S_Region_SBs_R")
                
            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = south_import["OpeningStock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break

        elif SBU_site_stocks == "F":
            south_import = data.parse("S_Region_SBs_PET")
                
            while True:
                try:
                    set_time = input("Please enter the number of seconds you need to navigate the cursor to the top of your blank space \n")
                    set_time = int(set_time)
                    break
                except ValueError:
                    print("Make sure you entered a number")

            Opening_stock = south_import["OpeningStock"].astype(str)

            for a, b in enumerate(Opening_stock):
                if a==0:
                    time.sleep(set_time)
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
                else:
                    keyboard.type(b)
                    keyboard.press(Key.down)
                    keyboard.release(Key.down)
            break

        else:
            x = input("Please make sure you entered either A, B, C, D, E or F. Press Enter to continue, and try again")

    quit_or_not = input("Thank you for using Barnton Exporter. Press q to quit. However, if you want to use Barnton Exporter again for another site, press y.")

    if quit_or_not == "y":
        continue
    else:
        break
