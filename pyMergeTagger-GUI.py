import argparse
import time
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
import sys
import os
import re
from PIL import Image, ImageQt
import zipfile
from io import StringIO, BytesIO
from lxml import html
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait

from PyQt6.QtCore import Qt, QObject, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QPushButton,
    QWidget,
    QGridLayout,
    QLineEdit,
    QHBoxLayout,
    QVBoxLayout,
    QListWidget,    
    QSpinBox,    
    QTextEdit,
    QProgressBar,    
    QComboBox,
    QAbstractItemView,
    QTreeWidget,
    QTreeWidgetItem,
    QListWidgetItem,
    QFileDialog,
    QMessageBox  
    
                
)
from PyQt6.QtGui import QIcon, QPixmap, QAction, QActionGroup, QMovie

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'realblack7.pyMergeTagger.v.1'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class MergerWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)     

    def __init__(self, booksToMerge, allMetaData, checkedMode):
        QThread.__init__(self)
        self.booksToMerge = booksToMerge 
        self.allMetaData = allMetaData
        self.checkedMode = checkedMode 

    def runMerger(self):
        """Edit Files: Merge, Rename or Add Metadata"""        
        
        for index in range(len(self.booksToMerge)):

            if self.checkedMode == 'add data && rename' or self.checkedMode == 'only add data' or self.checkedMode == 'merge && add data':
                                
                with StringIO() as f:
                    f.write('<?xml version="1.0"?>\n')
                    f.write('<ComicInfo xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n')
                    f.write('    <Series>'+self.allMetaData[1]+'</Series>\n')
                                
                    f.write("    <Volume>"+re.sub('.*?([0-9.]*)$',r'\1',self.booksToMerge[index].text(0))+"</Volume>\n")                          
                                
                    if self.checkedMode == 'merge && add data' or self.checkedMode == 'only merge files':
                        f.write('    <Title>'+self.booksToMerge[index].text(0)+'</Title>\n')

                    elif self.checkedMode == 'only add data':
                        f.write('    <Title>'+self.booksToMerge[index].text(0).split(' - ')[1]+'</Title>\n')
                    else:
                        f.write('    <Title>'+self.allMetaData[0]+' '+re.sub('.*?([0-9.]*)$',r'\1',self.booksToMerge[index].text(0))+'</Title>\n') # Chapter x

                    f.write('    <Writer>'+self.allMetaData[2]+'</Writer>\n')

                    f.write('    <Genre>'+self.allMetaData[4]+'</Genre>\n')

                    f.write('    <Summary>'+self.allMetaData[5]+'</Summary>\n')
                                                    
                    f.write('    <Year>'+self.allMetaData[3]+'</Year>\n')

                    if self.allMetaData[6] == 'Manga':
                        f.write('    <Manga>YesAndRightToLeft</Manga>\n')

                    f.write('</ComicInfo>')
                    comicInfo = f.getvalue()
            
            zipmode = 'w'
            
            if self.checkedMode == 'add data && rename' or self.checkedMode == 'only add data':
                check_src_zip = self.booksToMerge[index].data(0, 3)    
                with zipfile.ZipFile(check_src_zip, 'r', compression=zipfile.ZIP_DEFLATED) as check_src_zip_file:
                    if not 'ComicInfo.xml' in check_src_zip_file.namelist():
                        zipmode = 'a'
            
            if self.checkedMode == 'add data && rename' or self.checkedMode == 'only rename':

                if zipmode == 'w':
                    dest_zip = os.path.join(os.path.dirname(self.booksToMerge[index].data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.booksToMerge[index].text(0))+'.cbz')
                else:
                    dest_zip = self.booksToMerge[index].data(0, 3)
                    rename_zip = os.path.join(os.path.dirname(self.booksToMerge[index].data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.booksToMerge[index].text(0))+'.cbz')
            
            elif self.checkedMode == 'only add data':

                if zipmode == 'w':
                    dest_zip = os.path.join(os.path.dirname(self.booksToMerge[index].data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.booksToMerge[index].text(0))+'.tmp')
                    rename_zip = os.path.join(os.path.dirname(self.booksToMerge[index].data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.booksToMerge[index].text(0))+'.cbz')
                else:
                    dest_zip = self.booksToMerge[index].data(0, 3)
                    rename_zip = os.path.join(os.path.dirname(self.booksToMerge[index].data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.booksToMerge[index].text(0))+'.cbz')
            
            elif self.checkedMode == 'merge && add data' or self.checkedMode == 'only merge files':
                
                dest_zip = os.path.join(os.path.dirname(self.booksToMerge[index].child(0).data(0, 3)), self.allMetaData[1]+' - '+self.booksToMerge[index].text(0)+'.cbz')
            
                
            if  self.checkedMode == 'only rename':                               
             
                os.rename(self.booksToMerge[index].data(0, 3), dest_zip)

            else:
                with zipfile.ZipFile(dest_zip, zipmode, compression=zipfile.ZIP_STORED) as dest_zip_file:
                    if self.checkedMode == 'add data && rename' or self.checkedMode == 'only add data':  

                        if zipmode == 'w':           

                            src_zip = self.booksToMerge[index].data(0, 3)     ###path of zip to extract from                 
                                            
                            with zipfile.ZipFile(src_zip, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:                         
                                for zitem in src_zip_file.namelist():
                                    if zitem != 'ComicInfo.xml':                                
                                        dest_zip_file.writestr(zitem, src_zip_file.read(zitem))
                            
                            os.remove(src_zip)                       
                                                    

                    elif self.checkedMode == 'merge && add data' or self.checkedMode == 'only merge files': 
                        
                        for x in range(self.booksToMerge[index].childCount()):                        
                            
                            src_zip = self.booksToMerge[index].child(x).data(0, 3)
                            
                            with zipfile.ZipFile(src_zip, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file: 
                                for zitem in src_zip_file.namelist():
                                    if zitem != 'ComicInfo.xml':                                   
                                                            
                                        dest_zip_file.writestr(os.path.join(self.booksToMerge[index].child(x).text(0), zitem), src_zip_file.read(zitem))
                            os.remove(src_zip)                

                    
                    if self.checkedMode == 'add data && rename' or self.checkedMode == 'only add data' or self.checkedMode == 'merge && add data':    
                        dest_zip_file.writestr('ComicInfo.xml', comicInfo)

                if 'rename_zip' in locals():
                    try:      
                        os.rename(dest_zip, rename_zip) 
                    except:
                        os.rename(dest_zip, os.path.splitext(rename_zip)[0] + '_duplicate.cbz')                    
                    del rename_zip

            
            self.progress.emit(index)
        
        self.finished.emit()

class RemoverWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(object)     

    def __init__(self, filesToDelete, tempZipFilePath, zipFilePath):
        QThread.__init__(self)
        self.filesToDelete = filesToDelete    
        self.tempZipFilePath = tempZipFilePath    
        self.zipFilePath = zipFilePath


    def runDeletion(self):
        with zipfile.ZipFile(self.tempZipFilePath, 'w', compression=zipfile.ZIP_STORED) as dest_zip_file:                                
                        
            with zipfile.ZipFile(self.zipFilePath, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:

                filesToKeep = [item for item in src_zip_file.namelist() if item not in self.filesToDelete]  
                i = 0              
                for zitem in filesToKeep:                                                                        
                    dest_zip_file.writestr(zitem, src_zip_file.read(zitem))
                    self.progress.emit([len(filesToKeep), i])
                    i = i + 1
                        
            os.remove(self.zipFilePath) 

        os.rename(self.tempZipFilePath, self.zipFilePath)
        self.finished.emit()

class ImageRemoverWindow(QWidget):
    finished = pyqtSignal()

    def __init__(self, zipFilePath):
        super().__init__()
        self.zipFilePath = zipFilePath
        self.zipFileName = os.path.splitext(os.path.basename(self.zipFilePath))[0]
        self.closeMenu = True
        self.setWindowTitle('Delete pages from file')
        self._createGUI()
        self.listAllImagesInZip() 

    def _createGUI(self):    

        layout1 = QHBoxLayout()
        layout2 = QVBoxLayout()
        self.label = QLabel(self.zipFileName)

        self.List1Drop = QListWidget()      ### add ALL the files
        self.List1Drop.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.List1Drop.setFixedHeight(300)
        self.List1Drop.setFixedWidth(260)
        self.List1Drop.setSortingEnabled(True)
        #self.List1Drop.itemClicked.connect(lambda: self.ShowPicture(0))
        self.List1Drop.itemSelectionChanged.connect(lambda: self.ShowPicture(0))

        self.addImageButton = QPushButton('>>')   ### add selected Chapters to Volume
        self.addImageButton.setFixedWidth(50)        
        self.addImageButton.clicked.connect(self.addImagesToDeleteSelection)
        self.removeImageButton = QPushButton('<<') ### remove selected Chapters from Volume
        self.removeImageButton.setFixedWidth(50)
        self.removeImageButton.clicked.connect(self.addImagesToFileSelection)

        #self.List3Drop = QListWidget()      ### add only files to a book
        self.List3Drop = QListWidget()        
        self.List3Drop.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.List3Drop.setFixedHeight(300)
        self.List3Drop.setFixedWidth(260)
        self.List3Drop.setSortingEnabled(False) 
        self.List3Drop.itemSelectionChanged.connect(lambda: self.ShowPicture(1))       
        
        self.displayFirstImageOfZip = QLabel()
        self.displayFirstImageOfZip.setFixedHeight(210)
        self.displayFirstImageOfZip.setFixedWidth(150)
        self.displayFirstImageOfZip.setStyleSheet("border: 1px solid gray;")
        self.displayFirstImageOfZip.setScaledContents(True)

        self.Button1Name = QPushButton('Delete')
        self.Button1Name.setFixedWidth(60)
        self.Button1Name.clicked.connect(self.deleteFilesFromZip)        

        self.showProgressBar =QProgressBar()
        self.showProgressBar.setFixedWidth(673)

        subProgessLayout = QHBoxLayout()

        subProgessLayout.addWidget(self.showProgressBar)
        subProgessLayout.addWidget(self.Button1Name)
        subProgessLayout.addStretch()

        subButtonLayout = QVBoxLayout()        
                
        subButtonLayout.addWidget(self.addImageButton)
        subButtonLayout.addWidget(self.removeImageButton)
        subButtonLayout.addStretch() 

        #layout.addWidget(self.label)
        layout1.addWidget(self.displayFirstImageOfZip)
        layout1.addWidget(self.List1Drop)
        layout1.addLayout(subButtonLayout)
        layout1.addWidget(self.List3Drop) 
        

        layout2.addWidget(self.label)
        layout2.addLayout(layout1),
        layout2.addLayout(subProgessLayout)

        self.setLayout(layout2)

    def listAllImagesInZip(self):
        with zipfile.ZipFile(self.zipFilePath, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:
            zipList = sorted(src_zip_file.namelist())            
            for zitem in zipList:
                if (".jpg" in zitem or ".JPG" in zitem or ".png" in zitem or ".PNG" in zitem):
                    bookImage = QListWidgetItem(zitem)
                    self.List1Drop.addItem(bookImage)
    
    def ShowPicture(self, n):
        
        if n == 0:
            zitem = self.List1Drop.currentItem().text()                    
        else:
            zitem = self.List3Drop.currentItem().text()             
        

        with zipfile.ZipFile(self.zipFilePath, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:            
            
            data = src_zip_file.read(zitem)
            dataEnc = BytesIO(data)
            size = 150, 210
            with Image.open(dataEnc) as zimg: 
                zimg.thumbnail(size, Image.NEAREST)                       
                
                qt_image = ImageQt.ImageQt(zimg)                        
                self.displayFirstImageOfZip.setPixmap(QPixmap.fromImage(qt_image).copy())    

    def addImagesToDeleteSelection(self):        
        self.List1Drop.blockSignals(True)
        if len(self.List1Drop.selectedItems()) != 0:
            
            list = []
            for index in range(len(self.List1Drop)):
                item = self.List1Drop.item(index)                         
                
                if item.isSelected():                    
                    list.append(index)

            list.sort(reverse=True)                

            for x in list:                
                
                imageFileToDelte = QListWidgetItem(self.List1Drop.item(x).text())
                self.List3Drop.addItem(imageFileToDelte)         
                self.List1Drop.takeItem(x)               

        self.List1Drop.blockSignals(False)

    def addImagesToFileSelection(self):
        self.List3Drop.blockSignals(True)
        if len(self.List3Drop.selectedItems()) != 0:
            
            list = []
            for index in range(len(self.List3Drop)):
                item = self.List3Drop.item(index)                         
                
                if item.isSelected():                    
                    list.append(index)

            list.sort(reverse=True)                

            for x in list:                
                
                imageFileToDelte = QListWidgetItem(self.List3Drop.item(x).text())
                self.List1Drop.addItem(imageFileToDelte)         
                self.List3Drop.takeItem(x) 
        self.List3Drop.blockSignals(False)     

    def reportProgress(self, n):
        if n[0] > 0:
            percentageProgress = 100 / n[0]

            
            self.showProgressBar.setFormat("%p%")
            if int(percentageProgress*(n[1]+1)) <= 100:
                self.showProgressBar.setValue(int(percentageProgress*(n[1]+1)))       
                                
            else:
                self.showProgressBar.setValue(100)

    def deleteFilesFromZip(self):
        self.List1Drop.setDisabled(True)
        self.List3Drop.setDisabled(True)
        self.addImageButton.setDisabled(True)
        self.removeImageButton.setDisabled(True)
        self.Button1Name.setDisabled(True)
        self.closeMenu = False

        filesToDelete = []
        for x in range(self.List3Drop.count()):
            filesToDelete.append(self.List3Drop.item(x).text())

        tempZipFilePath = os.path.splitext(self.zipFilePath)[0]+'.tmp'

        self.removerthread = QThread()            
        self.removeworker = RemoverWorker(filesToDelete, tempZipFilePath, self.zipFilePath)
            
        self.removeworker.moveToThread(self.removerthread)
            
        self.removerthread.started.connect(self.removeworker.runDeletion)
        self.removeworker.progress.connect(self.reportProgress) 
        self.removeworker.finished.connect(self.endDeletion)       
        self.removeworker.finished.connect(self.removerthread.quit)
        self.removeworker.finished.connect(self.removeworker.deleteLater)
        self.removerthread.finished.connect(self.removerthread.deleteLater)            
            
        self.removerthread.start()        
        
    def endDeletion(self):       
        self.List3Drop.clear() 
        self.List1Drop.setDisabled(False)
        self.List3Drop.setDisabled(False)
        self.addImageButton.setDisabled(False)
        self.removeImageButton.setDisabled(False)
        self.Button1Name.setDisabled(False)
        self.closeMenu = True

    def closeEvent(self, event):
        
        if self.closeMenu == True:
            event.accept()
            self.finished.emit()
        else:
            event.ignore()

class InternetBrowser(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)

    threadOutputMangaSearch =pyqtSignal(object)
    threadOutputGetMetaData =pyqtSignal(object)      

    def __init__(self):
        QThread.__init__(self)
        self.opts = Options()
        self.opts.headless=True 
        self.opts.add_argument("--log fatal")
        geckopath = os.path.join(os.path.dirname(__file__), "geckodriver.exe")      
        self.driver = webdriver.Firefox(service=Service(geckopath), options=self.opts)
        
    
    def runBrowserManga4Life(self, url):
         
        self.driver.get(url)

        WebDriverWait(self.driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')        

        raw_html = html.fromstring(self.driver.page_source)        

        aname = raw_html.xpath('//a[@class="SeriesName ng-binding"]/text()')
        ahref = raw_html.xpath('//a[@class="SeriesName ng-binding"]/@href')

        data = [aname, ahref]        

        self.threadOutputMangaSearch.emit(data)

    def runBrowserMyAnimeList(self, url):
        self.driver.get(url)

        WebDriverWait(self.driver, 15).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')        

        raw_html = html.fromstring(self.driver.page_source)        

        aname = raw_html.xpath('//tr/td/a[@class="hoverinfo_trigger fw-b"]/strong/text()')
        ahref = raw_html.xpath('//tr/td/a[@class="hoverinfo_trigger fw-b"]/@href')

        data = [aname, ahref]        

        self.threadOutputMangaSearch.emit(data) 

    def runGetMetaData(self, url):
        
        self.driver.get(url)
        WebDriverWait(self.driver, 15).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
        raw_html = html.fromstring(self.driver.page_source)

        self.threadOutputGetMetaData.emit(raw_html)

    def closeBrowser(self):
        self.driver.quit()
        self.finished.emit()

class MainWindow(QMainWindow):
    sendURLtoThreadMyAnimeList = pyqtSignal(str)
    sendURLtoThreadManga4Life = pyqtSignal(str)
    sendURLtoThreadGetMetaData = pyqtSignal(str)
    closeBrowser = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.w = None
        self.closeMenu = True 
        self.lastSearch = ''

        self.setWindowTitle("pyMergeTagger")        
        self.setAcceptDrops(True)
        self._createMenu()      
        self._createNameInput()
        self._createDragandDrop() 
        self._createMeta() 
        self._createMaster()
        ##own thread 
        
        self.browserthread = QThread()            
        self.browseworker = InternetBrowser()
            
        self.browseworker.moveToThread(self.browserthread)

        self.sendURLtoThreadManga4Life.connect(self.browseworker.runBrowserManga4Life)
        self.sendURLtoThreadMyAnimeList.connect(self.browseworker.runBrowserMyAnimeList)
        self.sendURLtoThreadGetMetaData.connect(self.browseworker.runGetMetaData)

        self.closeBrowser.connect(self.browseworker.closeBrowser)  
        self.browseworker.threadOutputMangaSearch.connect(self.writeMangaInfoInGUI)
        self.browseworker.threadOutputGetMetaData.connect(self.writeMetaDataInGUI)    
        #self.browseworker.progress.connect(self.reportProgress)         
          
        self.browserthread.start()  
   
    def _createMenu(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        modeMenu = menubar.addMenu('&Modes')

        self.openFileDialog = QAction(QIcon(os.path.join(os.path.dirname(__file__), 'folder.png')), 'Open', self)        
        self.openFileDialog.triggered.connect(self.addBrowsedFiles)
        self.openFileDialog.setShortcut("Ctrl+F")

        self.clearAllFields = QAction(QIcon(os.path.join(os.path.dirname(__file__), 'trash.png')), 'Clear fields', self)        
        self.clearAllFields.triggered.connect(self.clearEveryField)
        self.clearAllFields.setShortcut("Ctrl+S")

        self.modeMenuGroup = QActionGroup(menubar)
        self.modeMenuGroup.ExclusionPolicy.Exclusive

        self.checkMode0 = QAction('merge && add data', self)
        self.checkMode0.setShortcut("Ctrl+1") 
        self.checkMode0.setCheckable(True)        
        self.checkMode0.setChecked(True)                
        self.checkMode1 = QAction('only merge files', self)
        self.checkMode1.setShortcut("Ctrl+2") 
        self.checkMode1.setCheckable(True)
        self.checkMode2 = QAction('add data && rename',self)
        self.checkMode2.setShortcut("Ctrl+3") 
        self.checkMode2.setCheckable(True)
        self.checkMode3 = QAction('only add data',self)
        self.checkMode3.setShortcut("Ctrl+4") 
        self.checkMode3.setCheckable(True)
        self.checkMode4 = QAction('only rename',self)
        self.checkMode4.setShortcut("Ctrl+5") 
        self.checkMode4.setCheckable(True)

        self.checkMergeFiles = True
        self.checkMetaData = True
        self.checkRename = True     

        fileMenu.addAction(self.openFileDialog)
        fileMenu.addAction(self.clearAllFields)

        modeMenu.addAction(self.checkMode0)
        modeMenu.addAction(self.checkMode1)
        modeMenu.addAction(self.checkMode2)       
        modeMenu.addAction(self.checkMode3)
        modeMenu.addAction(self.checkMode4)

        self.modeMenuGroup.addAction(self.checkMode0)
        self.modeMenuGroup.addAction(self.checkMode1)
        self.modeMenuGroup.addAction(self.checkMode2)       
        self.modeMenuGroup.addAction(self.checkMode3)
        self.modeMenuGroup.addAction(self.checkMode4)

        self.checkMode0.triggered.connect(self.changeMode)        
        self.checkMode1.triggered.connect(self.changeMode)
        self.checkMode2.triggered.connect(self.changeMode)        
        self.checkMode3.triggered.connect(self.changeMode)        
        self.checkMode4.triggered.connect(self.changeMode)      

    def _createNameInput(self):        
        
        layoutName = QVBoxLayout()
        sublayout1Name = QHBoxLayout()  #Edit Name & No        
        sublayout3Name = QHBoxLayout()  #Label Name & No
      

        Label1Name = QLabel('Output Filename:')
        Label1Name.setFixedHeight(15)
        Label1Name.setFixedWidth(250)
           
        self.Edit1Name = QLineEdit() ### Filename
        self.Edit1Name.setFixedWidth(250)      

        
        self.Edit2Name = QSpinBox() ### Volume No
        self.Edit2Name.setValue(1)
        self.Edit2Name.setFixedWidth(45)
        self.Edit2Name.setMaximum(9999)

        self.Edit3Name = QLineEdit('Volume')
        self.Edit3Name.setFixedWidth(55)                  

        sublayout1Name.addWidget(self.Edit1Name)
        sublayout1Name.addWidget(self.Edit3Name)
        sublayout1Name.addWidget(self.Edit2Name)
        sublayout1Name.addStretch()        

        sublayout3Name.addWidget(Label1Name)        
        sublayout3Name.addStretch()     

        layoutName.addLayout(sublayout3Name)
        layoutName.addLayout(sublayout1Name)
                         
        layoutName.addStretch()           
        
        self.widgetName = QWidget()
        self.widgetName.setLayout(layoutName)          
        
    def _createDragandDrop(self):
        ### Drop & Info Layout ###              
        
        self.List1Drop = QListWidget()      ### add ALL the files
        self.List1Drop.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.List1Drop.setFixedHeight(300)
        self.List1Drop.setFixedWidth(260)
        self.List1Drop.setSortingEnabled(True)
        self.List1Drop.itemSelectionChanged.connect(lambda: self.ShowPicture(0)) 
        self.List1Drop.itemDoubleClicked.connect(self.showImageRemover)      

        self.addToVolumeButton = QPushButton('>>')   ### add selected Chapters to Volume
        self.addToVolumeButton.setFixedWidth(50)        
        self.addToVolumeButton.clicked.connect(self.addChapterstoVolumeSelection)
        self.removeFromVolumeButton = QPushButton('<<') ### remove selected Chapters from Volume
        self.removeFromVolumeButton.setFixedWidth(50)
        self.removeFromVolumeButton.clicked.connect(self.removeChaptersFromVolumeSelection)

        #self.List3Drop = QListWidget()      ### add only files to a book
        self.List3Drop = QTreeWidget()
        self.List3Drop.setHeaderHidden(True)
        self.List3Drop.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.List3Drop.setFixedHeight(300)
        self.List3Drop.setFixedWidth(260)
        self.List3Drop.setSortingEnabled(False) 
        self.List3Drop.itemSelectionChanged.connect(lambda: self.ShowPicture(1))        
        
        self.displayFirstImageOfZip = QLabel()
        self.displayFirstImageOfZip.setFixedHeight(210)
        self.displayFirstImageOfZip.setFixedWidth(150)
        self.displayFirstImageOfZip.setStyleSheet("border: 1px solid gray;")
        self.displayFirstImageOfZip.setScaledContents(True)        

        self.Button1Name = QPushButton('Merge')
        self.Button1Name.setFixedWidth(60)
        self.Button1Name.clicked.connect(self.mergeCBZFiles)        

        self.showProgressBar =QProgressBar()
        self.showProgressBar.setFixedWidth(673)
        self.showProgressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)

        DropLayout = QVBoxLayout()
        subDropLayout = QHBoxLayout()
        subButtonLayout = QVBoxLayout()
        subProgessLayout = QHBoxLayout()
                
        subButtonLayout.addWidget(self.addToVolumeButton)
        subButtonLayout.addWidget(self.removeFromVolumeButton)
        subButtonLayout.addStretch()

        subDropLayout.addWidget(self.displayFirstImageOfZip)
        subDropLayout.addWidget(self.List1Drop)
        subDropLayout.addLayout(subButtonLayout)
        subDropLayout.addWidget(self.List3Drop) 
        subDropLayout.addStretch() 

        subProgessLayout.addWidget(self.showProgressBar)
        subProgessLayout.addWidget(self.Button1Name)
        subProgessLayout.addStretch()  

        
        DropLayout.addLayout(subDropLayout)  
        DropLayout.addLayout(subProgessLayout)                   
        DropLayout.addStretch()

        self.WidgetDrop = QWidget()
        self.WidgetDrop.setLayout(DropLayout)

    def _createMeta(self):
        ### META LAYOUT ###
        self.searchManga = QLineEdit()
        self.searchManga.returnPressed.connect(self.searchMangaData)
        self.searchManga.setPlaceholderText('Search...')
        self.searchManga.setFixedWidth(250)

        self.Edit1Meta = QComboBox() #name        
        self.Edit1Meta.setFixedWidth(250)
        self.Edit1Meta.currentTextChanged.connect(self.getMetaData)  

        self.ChooseProvider = QComboBox() #manga4life, myanimelist  
        self.ChooseProvider.addItems(['Manga4Life', 'MyAnimeList'])    
        self.ChooseProvider.setFixedWidth(250)
        self.ChooseProvider.currentTextChanged.connect(self.changeDataProvider)       

        self.Edit2Meta = QLineEdit() #Authors
        self.Edit2Meta.setFixedWidth(250)   
        self.Edit3Meta = QSpinBox() #Year
        self.Edit3Meta.setMaximum(9999)
        self.Edit3Meta.setSpecialValueText(' ')
        self.Edit3Meta.setFixedWidth(250)
        
        self.Edit4Meta = QLineEdit() #Genre
        self.Edit4Meta.setFixedWidth(250)
        self.Edit5Meta = QTextEdit() #Summary
        self.Edit5Meta.setFixedHeight(150)

        self.Edit6Meta = QComboBox() #Type
        self.Edit6Meta.addItems(['Manga', 'Comic'])      
        
        Label1Meta = QLabel('Titel:') #name
        Label1Meta.setFixedHeight(15)        
        Label2Meta = QLabel('Author(s):') #Authors
        Label2Meta.setFixedHeight(15)
        Label3Meta = QLabel('Year:') #Year
        Label3Meta.setFixedHeight(15)
        Label4Meta = QLabel('Genre(s):') #Genre
        Label4Meta.setFixedHeight(15)
        Label5Meta = QLabel('Summary:') #Summary
        Label5Meta.setFixedHeight(15)
        Label6Meta = QLabel('Type:') #Type
        Label6Meta.setFixedHeight(15)
        Label7Meta = QLabel('Provider:') #Provider
        Label7Meta.setFixedHeight(15)
        Label8Meta = QLabel('Search:') #Type
        Label8Meta.setFixedHeight(15)

        self.loadingGIF = QMovie(os.path.join(os.path.dirname(__file__), 'loading.gif'))
        self.loadingLabel = QLabel()
        self.loadingLabel.setFixedHeight(30)
        self.loadingLabel.setFixedWidth(30)        

        MetaLayout = QGridLayout()
        MetaLayout.addWidget(self.loadingLabel, 0, 0) 
        MetaLayout.addWidget(Label8Meta, 1, 0 ) 
        MetaLayout.addWidget(self.searchManga, 1, 1 ) 
        MetaLayout.addWidget(Label7Meta, 2, 0 )
        MetaLayout.addWidget(self.ChooseProvider, 2, 1)        
        MetaLayout.addWidget(Label1Meta, 3, 0 )
        MetaLayout.addWidget(self.Edit1Meta, 3, 1 )
        MetaLayout.addWidget(Label2Meta, 4, 0 )
        MetaLayout.addWidget(self.Edit2Meta, 4, 1 )
        MetaLayout.addWidget(Label3Meta, 5, 0 )
        MetaLayout.addWidget(self.Edit3Meta, 5, 1 )
        MetaLayout.addWidget(Label4Meta, 6, 0 )
        MetaLayout.addWidget(self.Edit4Meta, 6, 1 )
        MetaLayout.addWidget(Label6Meta, 7, 0 )
        MetaLayout.addWidget(self.Edit6Meta, 7, 1 )
        MetaLayout.addWidget(Label5Meta, 8, 0, 1, 2 )
        MetaLayout.addWidget(self.Edit5Meta, 9, 0, 1, 2 ) 
               
        
        MetaLayout.setRowStretch(MetaLayout.rowCount(), 1)        
        MetaLayout.setColumnStretch(MetaLayout.columnCount(), 1)

        self.widgetMeta = QWidget()
        self.widgetMeta.setLayout(MetaLayout)

    def _createMaster(self):
        ### Master Layout ###

        masterlayout = QGridLayout()       
        masterlayout.addWidget(self.widgetName,0 ,0)
        masterlayout.addWidget(self.WidgetDrop,1 ,0, 1, 1)
        masterlayout.addWidget(self.widgetMeta,0 ,1, 2, 1)

        masterWidget = QWidget()        
        masterWidget.setLayout(masterlayout)       

        self.setCentralWidget(masterWidget)            

    def searchMangaData(self):
    #multithread
        if self.searchManga.text() != '':
            self.loadingLabel.setMovie(self.loadingGIF)
            self.loadingGIF.start()
            if self.ChooseProvider.currentText() == 'Manga4Life':              

                url = 'http://manga4life.com/search/?sort=v&desc=true&name='+self.searchManga.text().replace(" ", "%20") 
                self.sendURLtoThreadManga4Life.emit(url)                      

            elif self.ChooseProvider.currentText() == 'MyAnimeList':
                url = 'https://myanimelist.net/manga.php?cat=manga&q='+self.searchManga.text().replace(" ", "%20")
                self.sendURLtoThreadMyAnimeList.emit(url)
                                       
    def writeMangaInfoInGUI(self, data):
        if self.searchManga.text() != '':
            if self.ChooseProvider.currentText() == 'Manga4Life':                           

                self.Edit1Meta.blockSignals(True)

                self.Edit1Meta.clear()

                for index in range(len(data[0])):            
                    self.Edit1Meta.addItem(data[0][index])                        
                    self.Edit1Meta.setItemData(index, 'http://manga4life.com'+data[1][index], 3)               
                self.Edit1Meta.blockSignals(False)            


            elif self.ChooseProvider.currentText() == 'MyAnimeList':                            

                self.Edit1Meta.blockSignals(True)

                self.Edit1Meta.clear()

                if len(data[0]) >= 10:
                    searchlength = 10

                for index in range(searchlength):            
                    self.Edit1Meta.addItem(data[0][index])                        
                    self.Edit1Meta.setItemData(index, data[1][index], 3)               
                self.Edit1Meta.blockSignals(False)
            
            self.loadingLabel.clear()
            self.loadingGIF.stop()
            self.getMetaData()

    def getMetaData(self):

    ##multithread        

        if self.checkMetaData:
            if len(self.Edit1Meta.currentText()) == 0:
                #self.searchMangaData()   
                return                          

            else:                
                self.loadingLabel.setMovie(self.loadingGIF)
                self.loadingGIF.start()
                url = self.Edit1Meta.itemData(self.Edit1Meta.currentIndex(), 3) ###input for workthread                
                self.sendURLtoThreadGetMetaData.emit(url)            

    def writeMetaDataInGUI(self, raw_html):
        if self.ChooseProvider.currentText() == 'Manga4Life':                           

            aauthors = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Author(s):"]/following-sibling::a/text()')
            agenre = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Genre(s):"]/following-sibling::a/text()')
            ayear = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Released:"]/following-sibling::a/text()')
            allsummary = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/div/text()')[0]

            allauthors = ''
            allgenre = ''

            for index in range(len(aauthors)):
                allauthors = allauthors + aauthors[index]
                if index != len(aauthors)-1:
                    allauthors = allauthors + ', '
                    

            for index in range(len(agenre)):
                allgenre = allgenre + agenre[index]
                if index != len(agenre)-1:
                    allgenre = allgenre + ', '                     

        elif self.ChooseProvider.currentText() == 'MyAnimeList':
            aauthors = raw_html.xpath('//div[@class="information-block di-ib clearfix"]/span[@class="information studio author"]/a/text()')
            agenre = raw_html.xpath('//div[@class="spaceit_pad"]/span[text() = "Genres:"]/following-sibling::a/text()')
            ayear = raw_html.xpath('//div[@class="spaceit_pad"]/span[text() = "Published:"]/following-sibling::text()[1]')
            asummary = raw_html.xpath('//td/span[@itemprop="description"]/descendant::text()')                                    

            allauthors = ''
            allgenre = ''
            allsummary = ''

            for index in range(len(aauthors)):                      
                allauthors = allauthors + str.upper(aauthors[index].split(',')[0]) + ' ' + aauthors[index].split(', ')[1]
                if index != len(aauthors)-1:
                    allauthors = allauthors + ', '
                    

            for index in range(len(agenre)):
                allgenre = allgenre + agenre[index].replace(" ", "")
                if index != len(agenre)-1:
                    allgenre = allgenre + ', '

            for index in range(len(asummary)):
                allsummary = allsummary + asummary[index] 

        ### output of workthread given to new function
        self.Edit2Meta.setText(allauthors)
        self.Edit4Meta.setText(allgenre)               
                                            
        self.Edit3Meta.setSpecialValueText(re.findall('(\d{4})',ayear[0])[0])
        self.Edit5Meta.setText(allsummary)                                                    
        self.Edit1Name.setText(self.Edit1Meta.currentText())
        self.loadingLabel.clear()
        self.loadingGIF.stop()

    def changeDataProvider(self):
        if self.searchManga.text() != '':
            self.searchMangaData()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            check = True
            for i in urls:
                i = i.toString()                
                ext = None
                if '.' in i:
                    ext = i.split('.')[-1]                

                if ext:
                    ext = ext.lower()
                    if not ext in ['cbz', 'zip']:
                        check = False
                        return
                else:
                    check = False
                    return                
                  
            if not check:
                event.ignore()
            else:
                event.accept()

            #event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]        
        self.mainFileDirectory = os.path.basename(os.path.split(files[0])[1])
          
        #if mode without metadata, fill edit1name instead 
        mangaName = ''
        searchNameSplit = self.mainFileDirectory.split(' - ')

        for length in range(len(searchNameSplit)-1):
            if length != 0:
                mangaName += ' - '            
            mangaName += searchNameSplit[length]             

        if self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only add data':
            self.searchManga.setText(mangaName)
        else:
            self.Edit1Name.setText(mangaName)
        
        lengthList = len(self.List1Drop)
        
        for parseFilePath in files:
            
            if not self.List1Drop.findItems(os.path.splitext(os.path.basename(parseFilePath))[0], Qt.MatchFlag.MatchCaseSensitive | Qt.MatchFlag.MatchFixedString):

                bookToAdd = QListWidgetItem(os.path.splitext(os.path.basename(parseFilePath))[0])                               
                bookToAdd.setData(3, parseFilePath)                
                self.List1Drop.addItem(bookToAdd)        
        
        
        if lengthList == 0 or self.lastSearch != mangaName:        
            self.searchMangaData()

        self.lastSearch = mangaName        
        
    def addChapterstoVolumeSelection(self):
        self.List1Drop.blockSignals(True)
        self.List3Drop.blockSignals(True)

        if len(self.List1Drop.selectedItems()) != 0:
            
            list = []
            for index in range(len(self.List1Drop)):
                item = self.List1Drop.item(index)                         
                
                if item.isSelected():                    
                    list.append(index)

            list.sort(reverse=True)  

            if self.checkMergeFiles:    
                #book1 = QTreeWidgetItem([self.Edit1Name.text() +' - '+ self.Edit3Name.text() +' '+ self.Edit2Name.text()])
                book1 = QTreeWidgetItem([self.Edit3Name.text() +' '+ self.Edit2Name.text()])
                book1.setFlags(book1.flags() | Qt.ItemFlag.ItemIsEditable)       
          
            
            for x in list:
                if self.checkMergeFiles:
                    bookChapter = QTreeWidgetItem([self.Edit1Name.text() +' - Chapter '+ re.sub('.*?([0-9.]*)$',r'\1',self.List1Drop.item(x).text())])
                else:
                    if self.modeMenuGroup.checkedAction().text() == 'only add data':
                        bookChapter = QTreeWidgetItem([os.path.splitext(os.path.basename(self.List1Drop.item(x).data(3)))[0]])
                    else:    
                        bookChapter = QTreeWidgetItem([self.Edit1Name.text() +' - '+ self.Edit3Name.text() +' '+ re.sub('.*?([0-9.]*)$',r'\1',self.List1Drop.item(x).text())])

                bookChapter.setData(0, 3, self.List1Drop.item(x).data(3))

                if self.checkMergeFiles:                    
                    book1.addChild(bookChapter) 

                else:              
                    
                    bookChapter.setFlags(bookChapter.flags() | Qt.ItemFlag.ItemIsEditable)  
                    self.List3Drop.addTopLevelItem(bookChapter)         
                self.List1Drop.takeItem(x)
                
            if self.checkMergeFiles:
                self.List3Drop.addTopLevelItem(book1)               
                self.Edit2Name.setValue(self.Edit2Name.value()+1)

        self.List1Drop.blockSignals(False) 
        self.List3Drop.blockSignals(False)

    def removeChaptersFromVolumeSelection(self):
        self.List3Drop.blockSignals(True)
        self.List1Drop.blockSignals(True)

        root = self.List3Drop.invisibleRootItem()
        for child in self.List3Drop.selectedItems():
            
            if not child.parent():

                if child.childCount() > 0:                  
                    
                    for x in range(child.childCount()):                        
                        
                        if type(child.child(x).data(0, 3)) is str:
                            if not self.List1Drop.findItems(os.path.splitext(os.path.basename(child.child(x).data(0, 3)))[0], Qt.MatchFlag.MatchCaseSensitive | Qt.MatchFlag.MatchFixedString):                    
                                bookChapter = QListWidgetItem(os.path.splitext(os.path.basename(child.child(x).data(0, 3)))[0])
                                bookChapter.setData(3, child.child(x).data(0, 3))
                                self.List1Drop.addItem(bookChapter)

                    root.removeChild(child)

                else:
                    
                    if type(child.data(0, 3)) is str:
                        if not self.List1Drop.findItems(os.path.splitext(os.path.basename(child.data(0, 3)))[0], Qt.MatchFlag.MatchCaseSensitive | Qt.MatchFlag.MatchFixedString):                    
                                bookChapter = QListWidgetItem(os.path.splitext(os.path.basename(child.data(0, 3)))[0])
                                bookChapter.setData(3, child.data(0, 3))
                                self.List1Drop.addItem(bookChapter)
                    root.removeChild(child)  

            else:
                ## if parent empty --> remove parent
                          
                lookIfEmpty = child.parent().childCount()-1
                parentOfChild = child.parent()            
                
                if not self.List1Drop.findItems(os.path.splitext(os.path.basename(child.data(0, 3)))[0], Qt.MatchFlag.MatchCaseSensitive | Qt.MatchFlag.MatchFixedString):
                    bookChapter = QListWidgetItem(os.path.splitext(os.path.basename(child.data(0, 3)))[0])
                    bookChapter.setData(3, child.data(0, 3))
                    self.List1Drop.addItem(bookChapter)

                (child.parent() or root).removeChild(child)
                if lookIfEmpty == 0:             
                    
                        root.removeChild(parentOfChild)
        self.List3Drop.blockSignals(False)
        self.List1Drop.blockSignals(False)
        
    def addBrowsedFiles(self):        
                
        filter = "Media (*.cbz *.zip)"
        fileName = QFileDialog.getOpenFileNames(None, 'Browse files', 'D:\Medien\Manga', filter)
        
        if fileName != ([], ''):
            self.mainFileDirectory = os.path.basename(fileName[0][0]) 
            mangaName = ''
            searchNameSplit = self.mainFileDirectory.split(' - ')

            for length in range(len(searchNameSplit)-1):
                if length != 0:
                    mangaName += ' - '            
                mangaName += searchNameSplit[length]          
                        
            if self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only add data':
                                
                self.searchManga.setText(mangaName)
            else:
                                
                self.Edit1Name.setText(mangaName)
            
            lengthList = len(self.List1Drop)
            for parseFilePath in fileName[0]:                
                              
                if not self.List1Drop.findItems(os.path.splitext(os.path.basename(parseFilePath))[0], Qt.MatchFlag.MatchCaseSensitive | Qt.MatchFlag.MatchFixedString):
                    bookToAdd = QListWidgetItem(os.path.splitext(os.path.basename(parseFilePath))[0])                                                                      
                    bookToAdd.setData(3, parseFilePath)                
                    self.List1Drop.addItem(bookToAdd)  

            if lengthList == 0 or self.lastSearch != mangaName:
                self.searchMangaData()
            
            self.lastSearch = mangaName

    def changeMode(self):
        addMetaStatePrior = self.checkMetaData

        if self.Edit1Name.text() == '':            
            self.Edit1Name.setText(self.searchManga.text())

        if self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only add data':
           self.checkMetaData = True           
        elif self.modeMenuGroup.checkedAction().text() == 'only merge files' or self.modeMenuGroup.checkedAction().text() == 'only rename':
            self.checkMetaData = False

        if self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'only merge files':
            self.checkMergeFiles = True
        elif self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only add data' or self.modeMenuGroup.checkedAction().text() == 'only rename':
            self.checkMergeFiles = False

        if self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only rename' or self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'only merge files':
           self.checkRename = True
        elif self.modeMenuGroup.checkedAction().text() == 'only add data':
            self.checkRename = False
        

        self.Edit1Meta.setDisabled(not self.checkMetaData)
        self.searchManga.setDisabled(not self.checkMetaData)
        self.Edit2Meta.setDisabled(not self.checkMetaData)
        self.Edit3Meta.setDisabled(not self.checkMetaData)
        self.Edit4Meta.setDisabled(not self.checkMetaData)
        self.Edit5Meta.setDisabled(not self.checkMetaData)
        self.Edit6Meta.setDisabled(not self.checkMetaData)
        self.ChooseProvider.setDisabled(not self.checkMetaData)

        self.Edit2Name.setDisabled(not self.checkMergeFiles)
        
        self.Edit1Name.setDisabled(not self.checkRename)
        self.Edit3Name.setDisabled(not self.checkRename)

        if self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only add data':
            self.Button1Name.setText('Add Data')
            
        elif self.modeMenuGroup.checkedAction().text() == 'only rename':
            self.Button1Name.setText('Rename')              
       
        else:
            self.Button1Name.setText('Merge')  

        if self.checkMetaData == True and addMetaStatePrior == False and len(self.Edit1Meta.currentText()) != 0:            
            self.searchMangaData()
        
    def reportProgress(self, n):
        
        if len(self.bookList) > 0:
            percentageProgress = 100 / len(self.bookList)

        if self.checkMergeFiles:
            self.showProgressBar.setFormat('Created '+self.List3Drop.topLevelItem(self.processed[n]).text(0)+" - %p%")
            if int(percentageProgress*(n+1)) <= 100:
                self.showProgressBar.setValue(int(percentageProgress*(n+1)))       
                               
            else:
                self.showProgressBar.setValue(100)          
                self.showProgressBar.setFormat("Finished - %p%")
        else:        
                 
            if self.checkMetaData:
                progText = 'Adding Data to'
            elif not self.checkMetaData:
                progText = 'Renaming'  
            ###if renaming only rename, no Zip actions
            #                
            self.showProgressBar.setFormat(progText+' '+self.List3Drop.topLevelItem(self.processed[n]).text(0)+" - %p%")
            if int(percentageProgress*(n+1)) <= 100:
                self.showProgressBar.setValue(int(percentageProgress*(n+1)))       
                               
            else:
                self.showProgressBar.setValue(100)          
                self.showProgressBar.setFormat("Finished - %p%")     
    
    def endMerging(self):     
        
        #remove only processed files, leave skipped files    

        for x in reversed(range(len(self.processed))):            
            self.List3Drop.takeTopLevelItem(self.processed[x])
                
        self.Button1Name.setDisabled(False)
        self.List1Drop.setDisabled(False)
        self.List3Drop.setDisabled(False)
        self.Edit1Meta.setDisabled(not self.checkMetaData)
        self.searchManga.setDisabled(not self.checkMetaData)
        self.Edit2Meta.setDisabled(not self.checkMetaData)
        self.Edit3Meta.setDisabled(not self.checkMetaData)
        self.Edit4Meta.setDisabled(not self.checkMetaData)
        self.Edit5Meta.setDisabled(not self.checkMetaData)
        self.Edit6Meta.setDisabled(not self.checkMetaData)
        self.Edit3Name.setDisabled(not self.checkRename)
        self.Edit2Name.setDisabled(not self.checkMergeFiles)
        self.Edit1Name.setDisabled(not self.checkRename) 
        self.ChooseProvider.setDisabled(not self.checkMetaData)
        self.displayFirstImageOfZip.clear()
        self.addToVolumeButton.setDisabled(False)
        self.removeFromVolumeButton.setDisabled(False)

        self.closeMenu = True 

    def mergeCBZFiles(self):
        self.Button1Name.setDisabled(True)
        self.List1Drop.setDisabled(True)
        self.List3Drop.setDisabled(True)
        self.Edit1Meta.setDisabled(True)
        self.searchManga.setDisabled(True)
        self.Edit2Meta.setDisabled(True)
        self.Edit3Meta.setDisabled(True)
        self.Edit4Meta.setDisabled(True)
        self.Edit5Meta.setDisabled(True)
        self.Edit6Meta.setDisabled(True)
        self.Edit3Name.setDisabled(True)
        self.Edit2Name.setDisabled(True)
        self.Edit1Name.setDisabled(True)
        self.ChooseProvider.setDisabled(True)
        self.addToVolumeButton.setDisabled(True)
        self.removeFromVolumeButton.setDisabled(True)

        self.closeMenu = False   

        if self.modeMenuGroup.checkedAction().text() == 'merge && add data' or self.modeMenuGroup.checkedAction().text() == 'add data && rename' or self.modeMenuGroup.checkedAction().text() == 'only merge files' or self.modeMenuGroup.checkedAction().text() == 'only rename':
            shouldCheckIfBookExists = True
        else:
            shouldCheckIfBookExists = False

        index = 0
        self.bookList = []
        self.processed = []
        self.skipped = []
        while index <= self.List3Drop.topLevelItemCount()-1:
            if shouldCheckIfBookExists:                
  
                if not self.checkMergeFiles:                                      
                        
                    if not os.path.exists(os.path.join(os.path.dirname(self.List3Drop.topLevelItem(index).data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.List3Drop.topLevelItem(index).text(0))+'.cbz')):
                            
                        self.processed.append(index)                        
                        self.bookList.append(self.List3Drop.topLevelItem(index))
                    else:                    
                        self.skipped.append(index)                     
                            
                else:
                                            
                    if not os.path.exists(os.path.join(os.path.dirname(self.List3Drop.topLevelItem(index).child(0).data(0, 3)), self.Edit1Name.text()+' - '+self.List3Drop.topLevelItem(index).text(0)+'.cbz')):
                            
                        self.processed.append(index)                             
                        self.bookList.append(self.List3Drop.topLevelItem(index)) 
                    else:                    
                        self.skipped.append(index) 

            else:
                self.bookList.append(self.List3Drop.topLevelItem(index))
                self.processed.append(index)
            index = index + 1

        if shouldCheckIfBookExists and len(self.skipped) != 0:    
            if not self.checkMergeFiles:
                    
                skippedDLG = QMessageBox(self)
                skippedDLG.setWindowTitle("Skip files?")
                textDLG = 'The following files already exist: \n\n'
                for y in range(len(self.skipped)):
                    textDLG += os.path.join(os.path.dirname(self.List3Drop.topLevelItem(self.skipped[y]).data(0, 3)), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', self.List3Drop.topLevelItem(self.skipped[y]).text(0))+'.cbz') + '\n'
                textDLG += '\n Continue without these files?'
                skippedDLG.setText(textDLG)
                skippedDLG.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                skippedDLG.setIcon(QMessageBox.Icon.Question)
                button = skippedDLG.exec()            
                        
            else:
                skippedDLG = QMessageBox(self)
                skippedDLG.setWindowTitle("Skip files?")
                textDLG = 'The following files already exist: \n\n'
                for y in range(len(self.skipped)):
                    textDLG += os.path.join(os.path.dirname(self.List3Drop.topLevelItem(self.skipped[y]).child(0).data(0, 3)), self.Edit1Name.text()+' - '+self.List3Drop.topLevelItem(self.skipped[y]).text(0)+'.cbz') + '\n'
                textDLG += '\n Continue without these files?'
                skippedDLG.setText(textDLG)
                skippedDLG.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                skippedDLG.setIcon(QMessageBox.Icon.Question)
                button = skippedDLG.exec()
        
        else:
            button = True

        if button == QMessageBox.StandardButton.Yes or button == True:

            metaDatas = [self.Edit3Name.text(), self.Edit1Name.text(), self.Edit2Meta.text(), self.Edit3Meta.text(), self.Edit4Meta.text(), self.Edit5Meta.toPlainText(), self.Edit6Meta.currentText()]

            checkMode= self.modeMenuGroup.checkedAction().text()         
            
            self.mergerthread = QThread()            
            self.mergeworker = MergerWorker(self.bookList, metaDatas, checkMode)
            
            self.mergeworker.moveToThread(self.mergerthread)
            
            self.mergerthread.started.connect(self.mergeworker.runMerger)
            self.mergeworker.progress.connect(self.reportProgress) 
            self.mergeworker.finished.connect(self.endMerging)       
            self.mergeworker.finished.connect(self.mergerthread.quit)
            self.mergeworker.finished.connect(self.mergeworker.deleteLater)
            self.mergerthread.finished.connect(self.mergerthread.deleteLater)            
            
            self.mergerthread.start()
        else:
            self.Button1Name.setDisabled(False)
            self.List1Drop.setDisabled(False)
            self.List3Drop.setDisabled(False)
            self.Edit1Meta.setDisabled(not self.checkMetaData)
            self.searchManga.setDisabled(not self.checkMetaData)
            self.Edit2Meta.setDisabled(not self.checkMetaData)
            self.Edit3Meta.setDisabled(not self.checkMetaData)
            self.Edit4Meta.setDisabled(not self.checkMetaData)
            self.Edit5Meta.setDisabled(not self.checkMetaData)
            self.Edit6Meta.setDisabled(not self.checkMetaData)
            self.Edit3Name.setDisabled(not self.checkRename)
            self.Edit2Name.setDisabled(not self.checkMergeFiles)
            self.Edit1Name.setDisabled(not self.checkRename)
            self.ChooseProvider.setDisabled(False)
            self.addToVolumeButton.setDisabled(False)
            self.removeFromVolumeButton.setDisabled(False)

            self.closeMenu = True       
                    
    def clearEveryField(self):           

        self.Edit1Name.clear()
        self.searchManga.clear()
        self.Edit1Meta.clear()
        self.Edit2Meta.clear()
        self.Edit4Meta.clear()
        self.Edit3Meta.setSpecialValueText(' ')
        self.Edit5Meta.clear()
        self.List1Drop.clear()
        self.List3Drop.clear()
        self.displayFirstImageOfZip.clear()
        self.Edit2Name.setValue(1) 
        self.showProgressBar.setValue(0)          
        self.showProgressBar.setFormat("")                             
    
    def showImageRemover(self):
        self.Button1Name.setDisabled(True)
        self.List1Drop.setDisabled(True)
        self.List3Drop.setDisabled(True)
        self.Edit1Meta.setDisabled(True)
        self.searchManga.setDisabled(True)
        self.Edit2Meta.setDisabled(True)
        self.Edit3Meta.setDisabled(True)
        self.Edit4Meta.setDisabled(True)
        self.Edit5Meta.setDisabled(True)
        self.Edit6Meta.setDisabled(True)
        self.Edit3Name.setDisabled(True)
        self.Edit2Name.setDisabled(True)
        self.Edit1Name.setDisabled(True)
        self.ChooseProvider.setDisabled(True)
        self.addToVolumeButton.setDisabled(True)
        self.removeFromVolumeButton.setDisabled(True)

        self.closeMenu = False

        if self.w is None:
            self.w = ImageRemoverWindow(self.List1Drop.currentItem().data(3))            
            self.w.show()
            self.w.finished.connect(self.closeImageRemover)            
        else:
            self.w.close()
            self.w = None     

    def closeImageRemover(self):
        self.w = None 
        self.Button1Name.setDisabled(False)
        self.List1Drop.setDisabled(False)
        self.List3Drop.setDisabled(False)
        self.Edit1Meta.setDisabled(not self.checkMetaData)
        self.searchManga.setDisabled(not self.checkMetaData)
        self.Edit2Meta.setDisabled(not self.checkMetaData)
        self.Edit3Meta.setDisabled(not self.checkMetaData)
        self.Edit4Meta.setDisabled(not self.checkMetaData)
        self.Edit5Meta.setDisabled(not self.checkMetaData)
        self.Edit6Meta.setDisabled(not self.checkMetaData)
        self.Edit3Name.setDisabled(not self.checkRename)
        self.Edit2Name.setDisabled(not self.checkMergeFiles)
        self.Edit1Name.setDisabled(not self.checkRename) 
        self.ChooseProvider.setDisabled(not self.checkMetaData)
        self.addToVolumeButton.setDisabled(False)
        self.removeFromVolumeButton.setDisabled(False)
        
        self.closeMenu = True          

    def ShowPicture(self, n):
        checked = False        
                
        if n == 0 and hasattr(self.List1Drop.currentItem(), 'data'):            
            imgInZip = self.List1Drop.currentItem().data(3)
            checked =True
                    
        elif n == 1 and hasattr(self.List3Drop.currentItem(), 'data'):

            if self.mergeCBZFiles:
                if self.List3Drop.currentItem().childCount() == 0:
                    imgInZip = self.List3Drop.currentItem().data(0, 3)
                else:
                    imgInZip = self.List3Drop.currentItem().child(0).data(0, 3)
            else:
                imgInZip = self.List3Drop.currentItem().data(0, 3)
            
            checked =True

        if checked:
            with zipfile.ZipFile(imgInZip, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:                   
                    
                zipList = sorted(src_zip_file.namelist())            
                for zitem in zipList:
                    if (".jpg" in zitem or ".JPG" in zitem or ".png" in zitem or ".PNG" in zitem):
                                
                        data = src_zip_file.read(zitem)
                        dataEnc = BytesIO(data)
                        size = 150, 210
                        with Image.open(dataEnc) as zimg: 
                            zimg.thumbnail(size, Image.NEAREST)                        
                            #zimg.show()
                            qt_image = ImageQt.ImageQt(zimg)                        
                            self.displayFirstImageOfZip.setPixmap(QPixmap.fromImage(qt_image).copy())
                    break    

    def closeEvent(self, event):
        # do stuff
       
        if self.closeMenu == True:
            event.accept()            
            #self.driver.quit()
            self.closeBrowser.emit()
            self.browseworker.finished.connect(self.browserthread.quit)
            self.browseworker.finished.connect(self.browseworker.deleteLater)
            self.browserthread.finished.connect(self.browserthread.deleteLater)
            
        else:
            event.ignore()

class HeadlessMonitor():

    def __init__(self):
        super().__init__()        

        args = parser.parse_args() 
        if os.path.isdir(args.path):
            self.path = args.path
        else:
            print('The given path is not valid!')
            sys.exit(0)
        
        global PAUSED
        PAUSED = False

        if args.mode == '0':
            self.checkedMode = 'add data & rename'
        elif args.mode == '1':
            self.checkedMode = 'only add data'
        elif args.mode == '2':
            self.checkedMode = 'only rename' 
        else:
            self.checkedMode = 'add data & rename'

        if args.provider == '0' :
            self.provider = 'Manga4Life'
        elif args.provider == '1': 
            self.provider = 'MyAnimeList'
        else:
            self.provider = 'Manga4Life'

        self.filename = re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', args.media)       

        print('Watching: ' + self.path +'\nMode: '+ self.checkedMode +'\nProvider: '+ self.provider +'\nNaming Scheme: %Manga% - %'+self.filename+'% %Number%.cbz')      

        if self.checkedMode == 'add data & rename' or self.checkedMode == 'only add data':
            ###start browser
            opts = Options()
            opts.headless=True 
            opts.add_argument("--log fatal")
            geckopath = os.path.join(os.path.dirname(__file__), "geckodriver.exe")      
            self.driver = webdriver.Firefox(service=Service(geckopath), options=opts) 
            print('\nBrowser started successfully\n')       

        self.lastMangaName = ''

        patterns = ["*.cbz", "*.zip"]
        ignore_patterns = None
        ignore_directories = False
        case_sensitive = True
        self.my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)
        print('Watchdog is ready.\nTo stop press "Ctrl" + "C"\n')
  
    def on_created(self, event):         
        global PAUSED
        # If PAUSED, exit
        if PAUSED is True:
            return

        # If not, pause anything else from running and continue
        PAUSED = True

        #time.sleep(5)
        isLocked = True
        while isLocked == True:
            try:
                with zipfile.ZipFile(event.src_path) as testfile:             
                    print('File ' + event.src_path + ' is okay')            
                
            except zipfile.BadZipFile:  
                time.sleep(2)                                          
                print('File ' + event.src_path + ' is not okay. Download may not be finished. Waiting...')                  
            else:
                # if mode = 0 or 1 get link and metaData
                mainFileDirectory = os.path.splitext(os.path.basename(event.src_path))[0]
                mangaName = ''
                searchNameSplit = mainFileDirectory.split(' - ')

                for length in range(len(searchNameSplit)-1):
                    if length != 0:
                        mangaName += ' - '            
                    mangaName += searchNameSplit[length]
                 
                volumeName = mainFileDirectory.split(' - ')[-1]                 

                #check if file already exists                                   
            
                if  self.checkedMode == 'only rename':  
                    newName = mangaName+ ' - ' + self.filename + ' '+ re.sub('.*?([0-9.]*)$',r'\1',volumeName)                              
                    if not os.path.exists(os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz')):                        
                        rename_zip = os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz')
                        os.rename(event.src_path, rename_zip)

                else:
                    
                    if self.lastMangaName != mangaName:                     
                        
                        if self.provider == 'Manga4Life':
                                        
                            url = 'http://manga4life.com/search/?sort=v&desc=true&name='+mangaName.replace(" ", "%20")                                       

                        elif self.provider == 'MyAnimeList':
                            url = 'https://myanimelist.net/manga.php?cat=manga&q='+mangaName.replace(" ", "%20")
                            
                        self.driver.get(url)
                        WebDriverWait(self.driver, 10).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')      
                        raw_html = html.fromstring(self.driver.page_source)   

                        if self.provider == 'Manga4Life':                                
                            aname = raw_html.xpath('//a[@class="SeriesName ng-binding"]/text()')
                            ahref = raw_html.xpath('//a[@class="SeriesName ng-binding"]/@href')
                        elif self.provider == 'MyAnimeList':
                            aname = raw_html.xpath('//tr/td/a[@class="hoverinfo_trigger fw-b"]/strong/text()')
                            ahref = raw_html.xpath('//tr/td/a[@class="hoverinfo_trigger fw-b"]/@href')

                        if self.provider == 'Manga4Life':                                
                            foundurl = 'http://manga4life.com'+ahref[0]                                     

                        elif self.provider == 'MyAnimeList':
                            foundurl = ahref[0]


                        data = [aname[0], foundurl]

                        self.driver.get(data[1])
                        WebDriverWait(self.driver, 15).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
                        raw_html = html.fromstring(self.driver.page_source)

                        if self.provider == 'Manga4Life':                         
                            aauthors = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Author(s):"]/following-sibling::a/text()')
                            agenre = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Genre(s):"]/following-sibling::a/text()')
                            ayear = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/span[text() = "Released:"]/following-sibling::a/text()')
                            allsummary = raw_html.xpath('//li[@class="list-group-item d-none d-md-block"]/div/text()')[0]

                            allauthors = ''
                            allgenre = ''

                            for index in range(len(aauthors)):
                                allauthors = allauthors + aauthors[index]
                                if index != len(aauthors)-1:
                                    allauthors = allauthors + ', '
                                                

                            for index in range(len(agenre)):
                                allgenre = allgenre + agenre[index]
                                if index != len(agenre)-1:
                                    allgenre = allgenre + ', '                     

                        elif self.provider == 'MyAnimeList':
                            aauthors = raw_html.xpath('//div[@class="information-block di-ib clearfix"]/span[@class="information studio author"]/a/text()')
                            agenre = raw_html.xpath('//div[@class="spaceit_pad"]/span[text() = "Genres:"]/following-sibling::a/text()')
                            ayear = raw_html.xpath('//div[@class="spaceit_pad"]/span[text() = "Published:"]/following-sibling::text()[1]')
                            asummary = raw_html.xpath('//td/span[@itemprop="description"]/descendant::text()')                                    

                            allauthors = ''
                            allgenre = ''
                            allsummary = ''

                            for index in range(len(aauthors)):                      
                                allauthors = allauthors + str.upper(aauthors[index].split(',')[0]) + ' ' + aauthors[index].split(', ')[1]
                                if index != len(aauthors)-1:
                                    allauthors = allauthors + ', '
                                                

                            for index in range(len(agenre)):
                                allgenre = allgenre + agenre[index].replace(" ", "")
                                if index != len(agenre)-1:
                                    allgenre = allgenre + ', '

                            for index in range(len(asummary)):
                                allsummary = allsummary + asummary[index]

                        self.allMetaData = [data[0], allauthors, allgenre, allsummary, re.findall('(\d{4})',ayear[0])[0]]                       
                    
                    name = 'ComicInfo.xml'                
                    with StringIO() as f:                                
                        f.write('<?xml version="1.0"?>\n')
                        f.write('<ComicInfo xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n')
                        f.write('    <Series>'+self.allMetaData[0]+'</Series>\n')
                                    
                        f.write("    <Volume>"+re.sub('.*?([0-9.]*)$',r'\1',volumeName)+"</Volume>\n")                        
                        
                        if self.checkedMode == 'only add data':
                            
                            f.write('    <Title>'+volumeName+'</Title>\n')
                        else:
                            f.write('    <Title>'+self.filename+' '+re.sub('.*?([0-9.]*)$',r'\1',volumeName)+'</Title>\n') # Chapter x

                        f.write('    <Writer>'+self.allMetaData[1]+'</Writer>\n')

                        f.write('    <Genre>'+self.allMetaData[2]+'</Genre>\n')

                        f.write('    <Summary>'+self.allMetaData[3]+'</Summary>\n')
                                                        
                        f.write('    <Year>'+self.allMetaData[4]+'</Year>\n')
                                
                        f.write('    <Manga>YesAndRightToLeft</Manga>\n')

                        f.write('</ComicInfo>')

                        content = f.getvalue()                                                                           
                     
                    doNothing = False 
                    doDelete = True
                    doRename = True
                    check_src_zip = event.src_path
                    with zipfile.ZipFile(check_src_zip, 'r', compression=zipfile.ZIP_DEFLATED) as check_src_zip_file:
                        if name in check_src_zip_file.namelist():

                            zipmode = 'w'
                            if self.checkedMode == 'add data & rename': 
                                newName = mangaName+ ' - ' + self.filename + ' '+ re.sub('.*?([0-9.]*)$',r'\1',volumeName)                                            

                                dest_zip = os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.tmp')

                                if event.src_path != os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz'):                                    
                                    rename_zip = os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz')
                                else:                                                                 
                                    doNothing = True                                                        
            
                            elif self.checkedMode == 'only add data':               
                                
                                dest_zip = os.path.splitext(event.src_path)[0]+'.tmp'
                                rename_zip = os.path.splitext(event.src_path)[0]+'.cbz'

                        else:                            
                            zipmode = 'a'
                            if self.checkedMode == 'add data & rename': 
                                newName = mangaName+ ' - ' + self.filename + ' '+ re.sub('.*?([0-9.]*)$',r'\1',volumeName)                                            

                                dest_zip = event.src_path 
                                
                                if event.src_path != os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz'):                                    
                                    doDelete = False
                                    rename_zip = os.path.join(os.path.dirname(event.src_path), re.sub(r'[\\/\:*"<>\|\%\$\^£]', '', newName)+'.cbz')                                     
                                else:                                                                                                           
                                    doDelete = False                                                                        
                                    doRename = False                                                
            
                            elif self.checkedMode == 'only add data':                                                                             
                                doDelete = False
                                dest_zip = event.src_path 

                    if not doNothing:
                        with zipfile.ZipFile(dest_zip, zipmode, compression=zipfile.ZIP_STORED) as dest_zip_file:                                     
                            
                            if zipmode == 'w':
                                src_zip = event.src_path     ###path of zip to extract from                 
                                                    
                                with zipfile.ZipFile(src_zip, 'r', compression=zipfile.ZIP_DEFLATED) as src_zip_file:                         
                                    for zitem in src_zip_file.namelist():
                                        if zitem != 'ComicInfo.xml':                                
                                            dest_zip_file.writestr(zitem, src_zip_file.read(zitem))
                                
                            dest_zip_file.writestr(name, content)
                            
                        if doDelete == True:                                                    
                            os.remove(src_zip) 

                        if doRename == True:                             
                            try:      
                                os.rename(dest_zip, rename_zip) 
                            except:
                                os.rename(dest_zip, os.path.splitext(rename_zip)[0] + '_duplicate.cbz')

                    if 'src_zip' in locals():
                        del src_zip 

                    if 'rename_zip' in locals():
                        del rename_zip

                    self.lastMangaName = mangaName
                    print('Processed '+event.src_path+' as '+self.allMetaData[0])

        
        # Once finished, allow other things to run
        PAUSED = False
       
    def watcher(self):
        self.my_event_handler.on_created = self.on_created
    
        go_recursively = True
        self.my_observer = Observer()    
        self.my_observer.schedule(self.my_event_handler, self.path, recursive=go_recursively)

        self.my_observer.start()

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.my_observer.stop()
            if self.checkedMode == 'add data & rename' or self.checkedMode == 'only add data':
                self.driver.quit()
        self.my_observer.join()

    def exit_handler(self):
        self.my_observer.stop()
        if self.checkedMode == 'add data & rename' or self.checkedMode == 'only add data':
            self.driver.quit()   
    
parser = argparse.ArgumentParser(description='Headless mode of pyMergeTagger.')
parser.add_argument('--path', metavar='C:/Manga', type=str, help='path of the monitored folder')
parser.add_argument('--mode', metavar='0', default='0', type=str, help='0: add data & rename, 1: only add data, 2: only rename. "0" by default.')
parser.add_argument('--provider', metavar='0', default='0', type=str, help='Provider for the Metadata."0" for Manga4Life, "1" for MyAnimeList. "0" by default')
parser.add_argument('--media', metavar='Chapter', default='Chapter', type=str, help='What media type the files are: e.g. Chapter, Book, Volume.\n Naming Scheme: %%Manga%% - %%media%%  %%Number%%.cbz. "Chapter" by default')

if parser.parse_args().path:
    runHeadlessMonitor = HeadlessMonitor()
    runHeadlessMonitor.watcher()    

else:                 
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(os.path.dirname(__file__), 'icon.ico')))
    window = MainWindow()
    window.show()
    app.processEvents()
    app.exec()
