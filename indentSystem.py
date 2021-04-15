# -*- coding: utf-8 -*
import os, pandas as pd
import traceback
import time
import xlwings as xw

from kivy.uix.boxlayout import BoxLayout
from kivy.app import App
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.core.window import Window
from kivy.factory import Factory
from kivy.uix.popup import Popup
from kivy.clock import Clock

import IndentController

# EXCELファイルの定義
wb = xw.Book('testPieceData.xlsm')
ws = wb.sheets['testPiece']

# Main view class
####################################################################
class MainRoot(BoxLayout):
    # Variant
    currentNumber = None
    startNumber   = None
    endNumber     = None
    remarksNumber = None
    endString     = ''
    outputText    = ''
    tempText      = ''

    # Initialize
    def __init__(self, **kwargs):
        self.maindisplay = Factory.Maindisplay()
        self.test = Factory.Test()
        super(MainRoot, self).__init__(**kwargs)
        Window.size = (300, 350)
    
    def change_disp0(self):
        self.clear_widgets()
        self.add_widget(self.maindisplay)
        self.maindisplay.ids.yy.focus = True

    # 入力補助
    def checkYYLen(self, text):
        if len(text) != 2: return
        self.maindisplay.ids.mm.focus = True
    
    def checkMMLen(self, text):
        if len(text) != 2: return
        self.maindisplay.ids.w.focus = True
    
    def checkWLen(self, text):
        if len(text) != 1: return
        self.maindisplay.ids.dd.focus = True
    
    # 呼び出し処理
    def callInformation(self):
        # Variant
        currentNumber = None
        startNumber   = None
        endNumber     = None
        remarksNumber = None
        endString     = ''
        outputText    = ''
        tempText      = ''

        if self.maindisplay.ids.yy.text == '' or self.maindisplay.ids.mm.text == '' :
            self.popupCaution = CausionPopup()
            self.popupCaution.open()
            self.popupCaution.ids.causion.text = 'No YY or MM!'
            return
        
        if ws.range((1,13)).value != None:
            self.currentNumber = int(ws.range((1,13)).value)
            self.startNumber   = int(ws.range((1,13)).value)
            self.endNumber     = int(ws.range((1,14)).value)
        if ws.range((1,15)).value != None: self.endString = ws.range((1,15)).value
        if ws.range((1,16)).value != None: self.remarksNumber = int(ws.range((1,16)).value)
        self.outputText    = ws.range((1,3)).value
        
        if self.remarksNumber != None and self.maindisplay.ids.w.text == '':
            self.popupCaution = CausionPopup()
            self.popupCaution.open()
            self.popupCaution.ids.causion.text = 'Please input W!'
            return
        
        # 出力形式に整形
        self.outputText = self.outputText.replace('YY', self.maindisplay.ids.yy.text)
        self.outputText = self.outputText.replace('MM', self.maindisplay.ids.mm.text)
        if self.remarksNumber == 2 or self.remarksNumber == 3:  
            self.outputText = self.outputText.replace('-', ' ' + self.maindisplay.ids.w.text + '-')       
        self.outputText = self.outputText.replace('DD', self.maindisplay.ids.dd.text)
        self.tempText   = self.outputText
        if ws.range((1,13)).value != None:
            self.outputText = self.outputText + str(self.startNumber) + str(self.endString)
        if self.remarksNumber == 1: self.outputText = self.outputText + self.maindisplay.ids.w.text        
        self.maindisplay.ids.outputText.text = self.outputText

    
    # 戻るボタン
    def backText(self):
        if self.currentNumber == self.startNumber: return
        if ws.range((1,13)).value == None: return
        self.currentNumber -= 1
        self.outputText = self.tempText + str(self.currentNumber) + str(self.endString)
        if self.remarksNumber == 1: self.outputText = self.outputText + self.maindisplay.ids.w.text        
        self.maindisplay.ids.outputText.text = self.outputText


    #進むボタン
    def proceedText(self):
        if self.currentNumber == self.endNumber: return
        if ws.range((1,13)).value == None: return
        self.currentNumber += 1
        self.outputText = self.tempText + str(self.currentNumber) + str(self.endString)
        if self.remarksNumber == 1: self.outputText = self.outputText + self.maindisplay.ids.w.text        
        self.maindisplay.ids.outputText.text = self.outputText
        
    
    # Send parts imformation to indenter
    def sendSignal(self):

        # Causion
        if self.maindisplay.ids.comPort.text == '':
            self.popupCaution = CausionPopup()
            self.popupCaution.open()
            self.popupCaution.ids.causion.text = 'NO COM PORT!'
            return
        if self.maindisplay.ids.outputText.text == '':
            self.popupCaution = CausionPopup()
            self.popupCaution.open()
            self.popupCaution.ids.causion.text = 'NO OUTPUT TEXT!'
            return
        
        # テキスト送信
        try:
            #Serial code sending
            self.outputText = "> " + self.outputText 
            contents = [self.outputText] # V0:text
            comPort = 'COM' + self.maindisplay.ids.comPort.text
            IndentController.sendSerial(contents, comPort)

            if ws.range((1,13)).value == None:
                self.popupFinish = FinishPopup()
                self.popupFinish.open()
                self.clearCells()
                return 

            # テキストの更新
            self.currentNumber += 1
            if self.currentNumber > self.endNumber:                    # 既定の数量を刻印完了時リセット
                self.popupFinish = FinishPopup()
                self.popupFinish.open()
                self.clearCells()
                return 

            self.outputText = self.tempText + str(self.currentNumber) + str(self.endString)
            if self.remarksNumber == 1: self.outputText = self.outputText + self.maindisplay.ids.w.text        
            self.maindisplay.ids.outputText.text = self.outputText

        except:
            traceback.print_exc()


    ## General functions       
    # Clear all interface cells
    def clearCells(self):
        self.maindisplay.ids.outputText.text = ''
        self.maindisplay.ids.yy.text         = ''
        self.maindisplay.ids.mm.text         = ''
        self.maindisplay.ids.w.text          = ''
        self.maindisplay.ids.dd.text         = ''
        self.currentNumber                   = 0
        self.startNumber                     = 0
        self.endNumber                       = 0
        self.endString                       = ''
        self.remarksNumber                   = None
        

# Popup class
####################################################################
class FinishPopup(Popup):
    def __init__(self, **kwargs):
        super(FinishPopup, self).__init__(**kwargs)
        # call dismiss_popup in 2 seconds
        Clock.schedule_once(self.dismiss_popup, 1)

    def dismiss_popup(self, dt):
        self.dismiss()

class CausionPopup(Popup):
    def __init__(self, **kwargs):
        super(CausionPopup, self).__init__(**kwargs)
        Clock.schedule_once(self.dismiss_popup, 2)

    def dismiss_popup(self, dt):
        self.dismiss()

# App class
####################################################################
class MainApp(App): 
    def __init__(self, **kwargs):
        super(MainApp, self).__init__(**kwargs)
        self.title = 'Indenter System'
    pass

if __name__ == "__main__":
    MainApp().run()