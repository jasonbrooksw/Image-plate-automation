import tifffile as tif
import numpy as np
import pywinauto
from pywinauto.application import Application
import os
import sys
import configparser
import ast
import datetime
import time
import easygui
import pygame
import subprocess

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
import string

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

def readLaunchFolder(ini, targetname):
    files = os.listdir(ini.savedirectory)
    try:
        [folder] = [elem for elem in files if targetname in elem and ini.shotnumber in elem and '.' not in elem]
        folderfile = os.listdir(ini.savedirectory+folder)
        [proctif] = [elem for elem in folderfile if targetname in elem and ini.shotnumber in elem and '.tif' in elem]
    except:
        print("Waiting for launch to finish...")
        t1 = datetime.datetime.now()
        t1 = t1.timestamp()
        while True:
            time.sleep(2)
            files = os.listdir(ini.savedirectory)
            t2 = datetime.datetime.now()
            t2 = t2.timestamp()
            if t2-t1>60 and t2-t1<=62:
                print("Launch appears to be taking a while")
                if ini.email != 'None':
                    send_mail(send_from = 'imageplatescan@gmail.com',
                    subject = "Launch appears to be taking a while - check scanner",
                    text = '',
                    send_to = ini.email,
                    password = ini.password,
                    files = None
                    )
            try:
                [folder] = [elem for elem in files if targetname in elem and ini.shotnumber in elem and '.' not in elem]
                folderfile = os.listdir(ini.savedirectory+folder)
                [proctif] = [elem for elem in folderfile if targetname in elem and ini.shotnumber in elem and '.tif' in elem]
                break
            except:
                continue
    print("found file created by launch")
    [rawgel] = [elem for elem in folderfile if targetname in elem and ini.shotnumber in elem and '.gel' in elem]
    proctif = ini.savedirectory+folder+'/'+proctif
    rawgel = ini.savedirectory+folder+'/'+rawgel
    return proctif, rawgel

# https://stackoverflow.com/questions/14894993/running-windows-shell-commands-with-python
def run_win_cmd(cmd):
    result = []
    process = subprocess.Popen(cmd,
                               shell=True,
                               stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE)
    for line in process.stdout:
        result.append(line)
    errcode = process.returncode
    print(cmd)
    print(line)
    if errcode is not None:
        raise Exception('cmd %s failed', cmd)

def copyTiff(ini, proctif):
    basepath = os.path.split(ini.savedirectory[:-1])[0]
    fullpath = basepath + '/organized IP data/'
    if ini.platename in ['SR 1', 'SR 2', 'MS 1', 'MS 2', 'CBS 1', 'CBS 2']:
        folder = fullpath+ini.platename[:-2]+'/'+ini.shotnumber+'/'
        if 'CBS' in ini.platename:
            fname = folder+ini.platename[-1]+'_'+str(ini.scannumber)+'.tiff'
        else:
            fname = folder+str(ini.scannumber)+'.tiff'
        cmd = 'copy ' + '"'+ proctif + '"' + ' ' + '"' + fname + '"'
        cmd = cmd.replace('/','\\')
        if not os.path.isdir(folder):
            os.makedirs(folder)
        run_win_cmd(cmd)

# https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
def send_mail(send_from: str, subject: str, text: str, 
send_to: list, password, files= None):
    default_address = 'imageplatescan@gmail.com'
    send_to= default_address if not send_to else send_to
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)  
    msg['Subject'] = subject
    msg.attach(MIMEText(text))
    for f in files or []:
        with open(f, "rb") as fil: 
            ext = f.split('.')[-1:]
            attachedfile = MIMEApplication(fil.read(), _subtype = ext)
            attachedfile.add_header(
                'content-disposition', 'attachment', filename = basename(f) )
        msg.attach(attachedfile)
    username = send_from
    smtp = smtplib.SMTP(host="smtp.gmail.com", port = 587) 
    smtp.starttls()
    smtp.login(username,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()
    
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
                 'SR 1':[[1000,1],[800,1],[600,1],[500,1]],
                 'MS 1':[[1000,1],[800,1],[600,1],[500,1]],
                 'SR 2':[[1000,1],[800,1],[600,1],[500,1]],
                 'MS 2':[[1000,1],[800,1],[600,1],[500,1]]}
        self.readVoltageVals(config)
        self.files = {}
        self.rereadconfig = False
        self.readDramaticMusic(config)
        self.email = config['MAIN']['emailAddress']
        if self.email != 'None':
            self.password = input("enter password for imageplatescan@gmail.com: ")
            
    def readDramaticMusic(self, config):
        '''this is the dumbest thing I have ever made'''
        val = config['MAIN']['dramaticMusic']
        self.dramaticmusic = []
        words = [item.replace('[','').replace(',','').replace(']','') for item in val]
        letters = [item for item in words if item]
        start = 0
        for i, l in enumerate(letters):
            if l.isdigit():
                self.dramaticmusic.append("".join(letters[start:i+1]))
                start = i+2
            i += 1
        
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
            for i,(voltage, freq) in enumerate(self.pmtvoltage):
                filefreq = len(vdict[str(voltage)])
                if filefreq < freq:
                    pmtv = str(voltage)
                    break
                elif voltage == self.pmtvoltage[-1][0]:
                    pmtv = str(voltage)
                    break
                elif filefreq == freq and str(self.pmtvoltage[i+1][0]) not in vdict.keys():
                    pmtv = str(self.pmtvoltage[i+1][0])
                    break
        else:
            pmtv = str(self.pmtvoltage[0][0])
        if int(pmtv) < 500 or int(pmtv)> 1000:
            print('assigned pmt voltage is not in allowed range')
            sys.exit()
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
        
    def sanityCheck(self):
        self.getComment()
        questiontext = ''
        SRtext = ''
        dialogbox = False
        if abs(int(self.comment[:-3]))>15:
            questiontext = '''    Are you sure you set the right time and shotnumber?
    It's been %s from shot time: %s; shot number is %s'''%(self.comment,self.shotTime,self.shotnumber)
            dialogbox = True
            
        files = np.array(os.listdir(self.savedirectory)) 
        prevshot = [fil for fil in files if self.shotnumber in fil]
        
        if prevshot:
            questiontext = '''    Are you sure you set the right time and shotnumber?
    It's been %s from shot time: %s; shot number is %s'''%(self.comment,self.shotTime,self.shotnumber)
            dialogbox = True
                
        SRplate = [item for item in self.plates if 'SR' in item]
        prevSR = []
        
        if SRplate:
            SRplate = SRplate[0]
            prevSR = [fil for fil in files if str(int(self.shotnumber)-2) in fil and SRplate in fil]
        
        if prevSR:
            SRtext = '''\n\n    It looks like you're using the same SR plate number (%s)
    as the last scan (previous shot number is %s) - is this okay?'''%(SRplate,str(int(self.shotnumber)-2))
            dialogbox = True
        if dialogbox:
            Q = easygui.ynbox(questiontext+SRtext+'\n\n'+"    If anything is wrong edit/save the .ini file and click reload ini file",
                  'Something wrong?', ("Reload ini file", 'Continue without reloading ini file'))
            if Q:
                self.rereadconfig = True

class click_time:
    '''does the clicky stuff'''
    
    def __init__(self, config):
        self.plates = dict(config['PLATES'])
        self.needednames = ['mQFileNameEdit','mQPMTEdit','mQCommentEdit']
    
    def getButtons(self, platename, scannumber):
        self.buttons = self.needednames[:]
        if scannumber == 1:
            for key, val in self.plates.items():
                if val == platename:
                    platenum = key[-1]
                    scale = self.plates['resscale' + platenum][:-2]
                    scale = 'p' + scale + 'Button'
                    scanreg = self.plates['scanregion' + platenum]
                    if 'grid' not in scanreg and 'free' not in scanreg:
                        scanreg = 'all' + scanreg + 'RadioButton'
                    elif 'grid' in scanreg:
                        scanreg = ['gridRadioButton', scanreg[-1]]
                    self.buttons.append(scale)
                    self.buttons.append(scanreg)
        
    def clickButtons(self, app, savename, comment, pmtv):
        dlg_spec = app.window(title='Typhoon FLA 7000', found_index=0)
        namedict = dict(zip(self.needednames,[savename, pmtv, comment]))
        for elem in self.buttons:
            time.sleep(1)
            if 'grid' in elem[0]:
                dlg_spec['all20x40RadioButton'].click()
                dlg_spec[elem].click()
                gridcol=dict(
                    zip(
                        list(string.ascii_uppercase[:16]),
                        sorted(range(16),reverse=True)
                    )
                )
                press_coords=(402,100)
                release_coords = (press_coords[0] - gridcol[elem[1]]*24,press_coords[1])
                dlg_spec['mQSamplingArea'].drag_mouse(press_coords=press_coords,
                                                   release_coords=release_coords)
            elif elem in self.needednames:
                dlg_spec[elem].click()
                dlg_spec[elem].type_keys('^a'+namedict[elem], with_spaces = True)
            else:
                dlg_spec[elem].click()
        dlg_spec['startButton'].click()
            
class monitor_scan: 
    
    def __init__(self):
        self.app = Application().connect(title = 'Typhoon FLA 7000')        
    
    def runScan(self, ini, click):
        ini.readPlateFiles()
        ini.getPmtValue()
        ini.getSaveName()
        ini.getComment()
        click.getButtons(ini.platename, ini.scannumber)
        click.clickButtons(self.app, ini.savename, ini.comment, ini.pmtv)
        if ini.scannumber == 1 and ini.platename in ini.dramaticmusic:
            pygame.mixer.init()
            pygame.mixer.music.load("Drool of Fats.mp3")
            pygame.mixer.music.play()
            pygame.mixer.music.fadeout(35*1000)
        
    def clickControl(self, ini, click):
        while True:
            files = os.listdir(ini.savedirectory)
            targetname = ini.platename + ' ' + str(ini.scannumber)
            plateloc=[elem for elem in files if targetname in elem and ini.shotnumber in elem and '.tif' in elem]
            if plateloc:
                time.sleep(1)
                dlg_spec = self.app.window(title='Typhoon FLA 7000', found_index = 0)
                dlg_spec['mQLaunchButton'].click()
                proctif, rawgel = readLaunchFolder(ini, targetname)
                copyTiff(ini, proctif)
                print(proctif)
                self.satQ = gel_op(rawgel).saturationCompare()
                second_window = self.app.window(title='Typhoon FLA 7000', found_index = 1)
                while second_window.exists():
                    dlg_spec = self.app.window(title='Typhoon FLA 7000', found_index = 0)
                    dlg_spec['mQReturnButton'].click()
                    time.sleep(2)
                    second_window = self.app.window(title='Typhoon FLA 7000', found_index = 1)
                time.sleep(1)
                if self.satQ:
                    self.runScan(ini, click)
                else:
                    if ini.email != 'None':
                        send_mail(send_from='imageplatescan@gmail.com',
                        subject=ini.shotnumber + ' ' + ini.platename + ' finished',
                        text='',
                        send_to= ini.email,
                        password = ini.password,
                        files= None
                        )
                    if ini.platename == ini.plates[-1]:
                        print('scanning is over')
                        pygame.mixer.init()
                        pygame.mixer.music.load("20th Century Fox Flute.mp3")
                        pygame.mixer.music.play()
                        time.sleep(30)
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
    ini.sanityCheck()
    while ini.rereadconfig:
        config.read(inifile)
        ini = ini_settings(config)
        ini.sanityCheck()
    click = click_time(config)
    mon = monitor_scan()
    mon.runScan(ini, click)
    mon.clickControl(ini,click)
    
run()