import tifffile as tif
import numpy as np
import pywinauto
import os
import sys
import configparser
import ast
import datetime
import time
import easygui
import pygame

# directory = 'D:/image_plate_test/'
# inifile = 'D:/image_plate_test/automate_image_plate.ini'
inifile = 'D:/Users/Scan/Documents/Users/LWFA/automate_image_plate files/automate_image_plate.ini'


def getFileNames(plates, directory, shotNumber):
    '''gets file names of plates with specified shot number'''
    
    files=np.array(os.listdir(directory))    
    file={}
    plateloc=[[j for j, elem in enumerate(files) if plates[i] in elem and shotNumber in elem and elem.find('.') == -1] for i in range(len(plates))]
    if plateloc:
        for i,coords in enumerate(plateloc):
            file[plates[i]]=[]
            if coords:
                for j in coords:
                    subfiles = np.array(os.listdir(directory+files[j]))
                    gelf = []
                    gelf = [f for f in subfiles if '.gel' in f]
                    if not gelf:
                        print("somethings wrong - couldn't find .gel file in folder " + files[j])
                        sys.exit()
                    file[plates[i]].append(directory+files[j]+'/'+gelf[0])
    return file
    
class ini_settings:
    '''does I/O shit'''
    
    def __init__(self, config):
        self.savedirectory = config['MAIN']['saveDirectory']
        self.shotnumber = config['MAIN']['shotNumber']
        self.date = config['MAIN']['date']
        self.shotTime = config['MAIN']['shotTime']
        self.plates = [value for key, value in dict(config['PLATES']).items() if 'plate' in key]
        self.platename = self.plates[0]
        self.defaults = {'CBS 1':[[1000,1],[900,1],[800,1],[700,1],[600,1]],
                 'CBS 2':[[1000,1],[900,1],[800,1],[700,1],[600,1]],
                 'SR 1':[[1000,1],[800,1],[600,1]],
                 'MS 1':[[1000,1],[800,1],[600,1]],
                 'SR 2':[[1000,1],[800,1],[600,1]],
                 'MS 2':[[1000,1],[800,1],[600,1]]}
        self.readVoltageVals(config)
        self.files = {}
        
    def readVoltageVals(self, config):
        voltlist = ast.literal_eval(config['MAIN']['pmtVoltage'])
        if voltlist:
            self.pmtvoltage = [[volt, freq] for volt, freq in voltlist if freq >=1]
            self.pmtvoltage = sorted(self.pmtvoltage, reverse = True)
        elif self.platename in self.defaults.keys():
            self.pmtvoltage = self.defaults[self.platename]
        else:
            print('check your pmt voltages you dingus')
            sys.exit()
        print('PMT voltage(s) to be used for this plate:')
        print(self.pmtvoltage)
        
    def readPlateFiles(self):
        self.files = getFileNames([self.platename], self.savedirectory, self.shotnumber)
        if self.files[self.platename]:
            self.scannumber = len(self.files[self.platename])+1
        else:
            self.scannumber = 1
        
    def getPmtValue(self):
        if self.files[self.platename]:
            vdict = {}
            for key, filenames in self.files.items():
                if key == self.platename:
                    for f in filenames:
                        pos = f.find('PMT')
                        v = f[pos + 3:pos + 7].replace('/', '')
                        if v in vdict.keys():
                            vdict[v].append(v)
                        else:
                            vdict[v] = [v]
                        
            for i, (voltage, freq) in enumerate(self.pmtvoltage):
                #this section isn't written too well - could probably be improved
                stop = 'bad' 
                if len(vdict[str(voltage)]) == freq and voltage != self.pmtvoltage[-1][0]:
                    for j in range(i+1, len(self.pmtvoltage)):
                            if str(self.pmtvoltage[j][0]) in vdict.keys():
                                if len(vdict[str(self.pmtvoltage[j][0])]) < freq:
                                    stop = j
                                    break
                            else:
                                stop = j
                                break
                    if stop != 'bad':
                        pmtv = str(self.pmtvoltage[stop][0])
                        break
                    else:
                        continue
                elif voltage == self.pmtvoltage[-1][0]:
                    pmtv = str(voltage)
                    break
                elif len(vdict[str(voltage)]) < freq:
                    pmtv = str(voltage)
        else:
            for voltage, freq in self.pmtvoltage:
                if freq > 0:
                    pmtv = str(voltage)
                    break
        self.pmtv = pmtv
        
    def getSaveName(self):
        s = self.shotnumber + ' ' + self.platename + ' ' + str(self.scannumber) + ' ' + 'PMT' + self.pmtv
        self.savename = s
        
    def getComment(self):
        t1 = datetime.datetime.strptime(self.date + ' ' + self.shotTime,'%Y-%m-%d %H:%M:%S')
        t1 = t1.timestamp()
        t2 = datetime.datetime.now()
        t2 = t2.timestamp()
        self.comment = str(int((t2-t1)/60)) + 'min'

class click_time:
    '''does the clicky stuff'''
    
    def __init__(self, config):
        self.buttons = {}
        self.plates = dict(config['PLATES'])
        for key, value in dict(config['BUTTONS']).items():
            self.buttons[key] = ast.literal_eval(value)
    
    def getUnusedButtons(self, platename, scannumber):
        needednames = ['filenamestart', 'comment', 'pmtsetting', 'runscan']
        resregionbuttons = []
        if scannumber == 1:
            for key, val in self.plates.items():
                if val == platename:
                    platenum = key[-1]
                    scale = self.plates['resscale' + platenum]
                    scanreg = self.plates['scanregion' + platenum]
                    resregionbuttons.append(scale)
                    if 'grid' not in scanreg:
                        resregionbuttons.append(scanreg)
                        break
                    else:
                        resregionbuttons.append('grid')
                        break
        self.unusedbuttons = [key for key in self.buttons.keys() if key not in needednames and key not in resregionbuttons]
        
    def clickButtons(self, savename, platename, typeables, specialnames = ['comment', 'pmtsetting']):
        'typeables = [comment, pmtv]'
        typedict = dict(zip(specialnames, typeables))
        self.usegrid = False
        for elem in self.buttons.keys():
            if elem not in self.unusedbuttons:
                time.sleep(1)
                pixcoord = self.buttons[elem]
                if elem in specialnames:
                    pywinauto.mouse.double_click(coords = pixcoord)
                    text = typedict[elem]
                    [pywinauto.keyboard.KeyAction(char).run() for char in text]
                elif elem == 'filenamestart':
                    pywinauto.mouse.press(coords = self.buttons['filenamestart'])
                    pywinauto.mouse.release(coords = self.buttons['filenameend'])
                    text = savename
                    [pywinauto.keyboard.KeyAction(char).run() for char in text]
                elif elem == 'grid':
                    self.usegrid = True
                    pywinauto.mouse.click(coords = self.buttons['fullregion'])
                    pywinauto.mouse.click(coords = pixcoord)
                    pywinauto.mouse.press(coords = self.buttons['gridstart'])
                    [p] = [key for key, value in self.plates.items() if value == platename]
                    num = self.plates['scanregion'+p[-1]][-1]
                    pywinauto.mouse.release(coords = self.buttons['gridend' + num])
                else:
                    pywinauto.mouse.click(coords = pixcoord)
                    
class monitor_scan: 
    
    def __init__(self):
        self.waittime = 65
    
    def runScan(self, ini, click):
        ini.readPlateFiles()
        ini.getPmtValue()
        ini.getSaveName()
        ini.getComment()
        click.getUnusedButtons(ini.platename, ini.scannumber)
        click.clickButtons(ini.savename, ini.platename, [ini.comment, ini.pmtv])
        
    def clickControl(self, ini, click):
        while True:
            files=os.listdir(ini.savedirectory)
            targetname = ini.platename + ' ' + str(ini.scannumber)
            plateloc=[elem for elem in files if targetname in elem and ini.shotnumber in elem and '.tif' in elem]
            if plateloc:
                time.sleep(1)
                pywinauto.mouse.click(coords = click.buttons['launch'])
                if click.usegrid:
                    self.waittime = 45
                else:
                    self.waittime = 65
                time.sleep(self.waittime)
                files=os.listdir(ini.savedirectory)
                [folder]=[elem for elem in files if targetname in elem and ini.shotnumber in elem and '.' not in elem]
                folderfile = os.listdir(ini.savedirectory+folder)
                [rawgel]=[elem for elem in folderfile if targetname in elem and ini.shotnumber in elem and '.gel' in elem]
                rawgel = ini.savedirectory+folder+'/'+rawgel
                print(rawgel)
                self.satQ = gel_op(rawgel).saturationCompare()
                pywinauto.mouse.click(coords = click.buttons['return'])
                time.sleep(1)
                if self.satQ:
                    self.runScan(ini, click)
                else:
                    if ini.platename == ini.plates[-1]:
                        print('scanning is over')
                        pygame.mixer.init()
                        pygame.mixer.music.load("20th Century Fox Flute.mp3")
                        pygame.mixer.music.play()
                        break
                    else:
                        [i] = [i for i, elem in enumerate(ini.plates) if elem == ini.platename]
                        Q = easygui.ynbox(ini.platename + ' finished. Click OK when ready to scan next plate: ' + ini.plates[i+1],
                                      'Scan Finished', ('OK', 'Cancel'))
                        if not Q:
                            sys.exit()
                        config = configparser.ConfigParser()
                        config.read(inifile)
                        ini.platename = ini.plates[i+1]
                        ini.readVoltageVals(config)
                        self.runScan(ini, click)
            time.sleep(2)

class gel_op:
    '''gel file operations class'''
    
    def __init__(self,platefile):
        self.gelfile = tif.TiffFile(platefile)
        
    def saturationCompare(self):
        gelarr = self.gelfile.asarray()
        gelmeta = self.gelfile.mdgel_metadata
        satval = np.max(gelmeta['ColorTable'])
        arrmax = np.max(gelarr)
        self.gelfile.close()
        if arrmax == satval:
            return True
        else:
            return False
        
def run():
    config = configparser.ConfigParser()
    config.read(inifile)
    ini = ini_settings(config)
    click = click_time(config)
    mon = monitor_scan()
    mon.runScan(ini, click)
    mon.clickControl(ini,click)
    
run()