global driver
global username , password
##########################################################################################
#                              Includes                                                  #
##########################################################################################
import importlib
import PyQt5.QtWidgets
from bs4 import Tag, NavigableString, Comment
from lxml import etree
from bs4 import BeautifulSoup
from PyQt5 import QtCore, QtGui, QtWidgets
from front import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import csv
import shutil
import datetime
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from PyQt5.QtGui import QTextCursor
import PyQt5.QtWidgets
from PyQt5.QtWidgets import QLineEdit
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import QtCore, QtGui, QtWidgets, QtTest
import sys
import os
import re
import pyexcel_xlsx
from pyexcel.cookbook import merge_all_to_a_book
import glob
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from PIL import Image
from PyQt5.QtGui import QTextCursor
from msedge.selenium_tools import Edge, EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.alert import Alert
import numpy as np
from tabulate import tabulate
import datetime
import winsound
import pandas as pd
from openpyxl import workbook, load_workbook
import openpyxl
import traceback
##########################################################################################
#                              Functions                                                 #
##########################################################################################
def click_x_path(xpath):
    varr = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    varr.click()

def data_to_x_path(xpath, data):
    varr = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    varr.send_keys(data)

def get_xpath(element):
    components = []
    child = element if element.name else element.parent
    for parent in child.parents:
        siblings = parent.find_all(child.name, recursive=False)
        components.append(
            child.name
            if siblings == [child] else
            '%s[%d]' % (child.name, 1 + siblings.index(child))
            )
        child = parent
    components.reverse()
    return '/%s' % '/'.join(components)

def wait( type, x, y, element_no):
    global recover_flag
    global driver
    counter = 0
    while True:
        if counter == 150:
            os.system('cls')
            print("recovering")
            recover_flag = True
            return
        os.system('cls')
        print(counter)
        time.sleep(.5)
        try:
            html = driver.page_source
            soup = BeautifulSoup(html, features="html.parser")
            elem = soup.find_all(type, {x: y})
            if (len(elem) > 0):
                break
            counter = counter + 1
        except:
            os.system('cls')
            print("recovering")
            recover_flag = True
            return
    return elem[element_no]


class EXCEL_TO_DBC(QtWidgets.QWidget, Ui_MainWindow):
    global driver
    global username, password
    ##########################################################################################
    #   (Done)                     Global Variables                                          #
    ##########################################################################################

    def __init__(self):
        ##########################################################################################
        ##########################################################################################
        ##########################################################################################
        QtWidgets.QWidget.__init__(self)

        self.setupUi(MainWindow)
        self.Exit_BTN.clicked.connect(self.exit_func)
        self.Browse_BTN.clicked.connect(self.BUTTON_BROWSE_func)
        self.online_check_BTN.clicked.connect(self.online_check_func)
        self.lunch_saved_data_BTN.clicked.connect(self.open_saved_file)

    def open_saved_file(self):

        try:
            with open('Polarion_login.txt', 'r') as text_file:
                file = text_file.readlines()
                if len(file) > 0:
                    u = file[0].replace(" ", "")
                    p = file[1].replace(" ", "")
                    input_username = u
                    input_password = p
                else:
                    QMessageBox.about(self, "Message", "Login file is empty!")
                    return
        except:
            QMessageBox.about(self, "Message", "Login file was not found")
            return


        self.Open_Polarion(input_username, input_password)


    '''
    Input: this function takes polarion username and password  
    Decreption: it opens polarion and download can messges excel file 
    Output: no output
    '''
    def Open_Polarion(self , user, pas):
        self.textBrowser.setText("Opening Polarion")
        global driver

        try:
            driver = webdriver.Edge(executable_path="Edge.exe")
        except:
            driver = Edge(executable_path=EdgeChromiumDriverManager().install())

        if len(user) == 0:
            print("Empty user")
            exit()
        if len(pas) == 0:
            print("Empty pas")
            exit()

        driver.get("https://alm.mahle/polarion")
        driver.maximize_window()

        user = user.replace('\n', "")
        data_to_x_path('//*[@id="j_username"]', user)  ## username
        data_to_x_path('//*[@id="j_password"]', pas)  ## password
        click_x_path('//*[@id="submitButton"]')  ## login

        self.textBrowser.setText("Logged in successfully")


    def Excel_to_DBC(self):
        pass
        # Excel to DBC Code


    '''
    Input: no input
    Description: it 
    Output: no output
    '''
    def BUTTON_BROWSE_func(self):

            global ext
            global file_name
            file_name = QFileDialog.getOpenFileName(self, 'Open File', 'Select File', '(*.xlsx)')
            try:
                file_name = file_name[0]
                only_file_name = os.path.basename(file_name)
                if only_file_name.endswith(".xlsx"):
                    pass
                else:
                    QMessageBox.about(self, "Message", "Please select Excel file type")

                self.textEdit_2.setText(only_file_name)







            except:
                pass
    def online_check_func(self):
        input_username = self.lineEdit_2.text()
        input_password = self.lineEdit.text()
        input_username = input_username.replace(" ", "")
        input_password = input_password.replace(" ", "")

        if self.lineEdit.text() == "" or self.lineEdit_2.text() == "" :
            QMessageBox.about(self, "Message", "fill your credentials first please")
            return
        else:
            pass

            ##########################################################################################
            #                          if input is empty                                             #
            ##########################################################################################
            if len(input_username) == 0 or len(input_password) == 0:
                reply = QMessageBox.question(self, 'Question', 'Use previously registered username and password ?',
                                             QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    ##########################################################################################
                    #                         check if there is saved login info                             #
                    ##########################################################################################
                    try:
                        with open('Polarion_login.txt', 'r') as text_file:
                            file = text_file.readlines()
                            if len(file) > 0:
                                u = file[0].replace(" ", "")
                                p = file[1].replace(" ", "")
                                input_username = u
                                input_password = p
                            else:
                                QMessageBox.about(self, "Message", "Login file is empty!")
                                self.finilization_code()
                    except:
                        QMessageBox.about(self, "Message", "Login file was not found")
                        self.finilization_code()
                else:
                    QMessageBox.about(self, "Message", "Thank You for using our tool")
                    self.finilization_code()
            ##########################################################################################
            #                          if checkbox is Checked                                        #
            ##########################################################################################
            if self.checkBox.isChecked() == True:
                print("open text file")
                with open('Polarion_login.txt', 'w') as text_file:
                    text_file.write(input_username)
                    text_file.write('\n')
                    text_file.write(input_password)
            else:
                print("remember me is not checked")


            self.Open_Polarion(input_username, input_password)

    def reflect_message(self, message):
        QMessageBox.about(self, "Message", str(message))

    def reflect_status(self , string):
        self.textBrowser.clear()
        self.textBrowser.append(str(string))


    def finilization_code(self):
        try:
            driver.close()
            quit()
        except:
            quit()


    def exit_func(self):
        QMessageBox.about(self, "Message", "Thank You for using the tool")
        self.finilization_code()





if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = EXCEL_TO_DBC()
    MainWindow.show()
    sys.exit(app.exec_())