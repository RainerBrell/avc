# -*- coding: UTF-8 -*-

"""
 Audio Video Converter for NVDA 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2022/2024 Rainer Brell nvda@brell.net 

 *** 22. Nov 2022, Version 1.0
 NVDA+y         Converts to MP3 
 NVDA+shift+y   Converts to MP4
 NVDA+Control+y opens the convert result folder 
 
 *** 19. Februar 2023, Version 1.1
 Works under NVDA 2023.1 
 localized to german 
 Progress sound changed from beep to silent wave sound
  *** the sound output can be deactivated on the NVDA Settings 
 write now a description	file with file extension .description	 
 *** can be deactivated on the NVDA Settings 
 Youtube result filename now safed without checksum 
 
 *** 2024.04.27 
 - works with NVDA 2024.1 
 - Works again with the current Firefox version
 - Customizied result folder 
 - A link with a multimedia extension is recognized and saved as an MP3 or MP4
 - the ffmpeg.exe is unziped on installation 
 - Translations into Ukrainian, Turkish, vietnames and more 
 
 *** 2024.06.24 
 - New yt-dlp.exe 
 
  *** 2024.11.01
  - New yt-dlp.exe 
  - Arabic localization
  - speakOnDemand = True 
 
 *** In Progress for future  *** 
 - The conversion process can be killed 
 - The Youtube downloader working in the background is automatically updated every 30 days 
 - convertion function into  windows explorer 
 
 """

import globalPluginHandler
from scriptHandler import script
import ui
import urllib 
import signal 
import gui
from gui.settingsDialogs import NVDASettingsDialog, SettingsPanel
from gui import guiHelper
import wx 
import tones 
import config 
import datetime 
import time 
import api
import os 
import nvwave 
import controlTypes
import scriptHandler
import subprocess
import threading
import platform
import addonHandler

addonHandler.initTranslation()

AddOnSummary     = addonHandler.getCodeAddon().manifest['summary']
AddOnName        = addonHandler.getCodeAddon().manifest['name']
AddOnPath        = os.path.dirname(__file__)
ToolsPath        = AddOnPath + "\\Tools\\"
SoundPath        = AddOnPath + "\\sounds\\"
AppData          = os.environ["APPDATA"]
DownloadPath     = AppData + "\\AVC-Results"
LogFile          = DownloadPath + "\\log.txt"
YouTubeVideo     = DownloadPath + "\\YouTubeVideo\\"
YouTubeAudio     = DownloadPath + "\\YouTubeAudio\\"
OtherVideo       = DownloadPath + "\\Video\\"
OtherAudio       = DownloadPath + "\\Audio\\"
YouTubeEXE       = AddOnPath + "\\Tools\\yt-dlp.exe"
ConverterEXE     = AddOnPath + "\\Tools\\ffmpeg.exe"
sectionName      = AddOnName 
processID        = None 

MultimediaExtensions= {
	"aac", 
	"avi", 
	"flac", 
	"mkv", 
	"m3u8", 
	"m4a", 
	"m4s", 
	"m4v", 
	"mpg", 
	"mov", 
	"mp2", 
	"mp3", 
	"mp4", 
	"mpeg", 
	"mpegts", 
	"ogg", 
	"ogv", 
	"oga", 
	"ts", 
	"vob", 
	"wav", 
	"webm", 
	"wmv", 
	"f4v", 
	"flv", 
	"swf", 
	"avchd", 
	"3gp"
}

invalidCharactersForFilename = [
	"/",
	"\\",
	":",
	"*",
	"<",
	">",
	"?",
	"!",
	"|"
]

def initConfiguration():
	confspec = { 
		"BeepWhileConverting":          "boolean(default=True)",
		"CheckYouTubeDownloaderUpdate": "boolean(default=True)",
		"ResultFolder":                 "string(default='')",
		"YouTubeDescription":           "boolean(default=True)",
		"Logging":                      "boolean(default=False)"
	}
	config.conf.spec[sectionName] = confspec

initConfiguration()

def getINI(key):
	"""  get nvda.ini value """ 
	value = config.conf[sectionName][key]
	log("Get INI Key: " + sectionName + " " + key + " " + str(value))
	return value 

def setINI(key, value):
	"""  set nvda.ini value """ 
	try:
		config.conf[sectionName][key] = value
		log("Set INI Key: " + str(value))
	except:
		log("Error in setINI")
		
def log(s):
	""" 
		Write to log file 
	""" 
	if config.conf[sectionName]["Logging"]:
		try:
			s = makePrintable(s)
			log = open(LogFile, "a")
			log.write(str(s))
			log.write("\n")
			log.close()
		except:
			pass

def PlayWave(filename):
	path = SoundPath + filename
	filename = path + ".wav"
	if os.path.exists(filename):
		if config.conf[sectionName]["BeepWhileConverting"]:
			try:
				#log("Play WAVE: " + filename)
				nvwave.playWaveFile(filename) 
			except:
				log("can not play wave file - " + filename)
	else:
		log("WAVE file not exist: " + filename)
		
def createFolder(Folder): 
	if not os.path.isdir(Folder):
		try:
			os.mkdir(Folder)
			log("Create Folder: " + Folder)
		except:
			# Translators: Could not create  folder.
			ui.message(_("Error - can not create folder"))
			log("Can not create folder: " + Folder)

def CheckFolders():
	global DownloadPath
	global LogFile
	global YouTubeVideo
	global YouTubeAudio
	global OtherVideo
	global OtherAudio
	# set new global vars 
	DownloadPath     = getINI("ResultFolder")
	LogFile          = DownloadPath + "\\log.txt"
	YouTubeVideo     = DownloadPath + "\\YouTubeVideo\\"
	YouTubeAudio     = DownloadPath + "\\YouTubeAudio\\"
	OtherVideo       = DownloadPath + "\\Video\\"
	OtherAudio       = DownloadPath + "\\Audio\\"
	# create folders 
	createFolder(DownloadPath)
	createFolder(YouTubeVideo)
	createFolder(YouTubeAudio)
	createFolder(OtherAudio)
	createFolder(OtherVideo)

def getCurrentAppName():
	Name = "Emty"
	try: 
		obj  = api.getForegroundObject()
		Name = str(obj.appModule.appName)
	except:
		Name = "Error AppName"
	return Name 

def getTime():
	return datetime.datetime.now().strftime("%H:%M:%S")

def getDate():
	return datetime.datetime.now().strftime("%d. %B %Y")
	
def signal():
	tones.beep(900, 200)
	
def isBrowser():
	"""
	 Verifies that NVDA is in a browser.
	""" 
	obj = api.getFocusObject()
	if not obj.treeInterceptor:
		# Translators: The user is not in a browser.
		ui.message(_("No browser window found"))
		log("Browser is not activ - exit")
		return False 
	else:
		return True

def getCurrentDocumentURL():
	""" 
		Get current 		unmasked document URL 
	"""
	URL = None 
	obj = api.getFocusObject()
	try:
		URL = obj.treeInterceptor.documentConstantIdentifier
		log("Masked  : " + URL)
		URL = urllib.parse.unquote(URL)
		log("Unmasked: " + URL)
	except:
		log("URL not found")
		return None 
	return URL 

def getLinkURL():
	""" 
	Get the current unmasked URL from the navigator objekt 
	"""
	obj = api.getNavigatorObject()
	URL = ""
	try:
		if obj.role == 19: #link 
			URL = obj.value 
			log("Masked Link URL  : " + URL) 
			URL = urllib.parse.unquote(URL)
			l = len(URL)
			# if last charakter "/" then kill it
			if URL[l-1: l] == "/":
				URL = URL[0: l-1]
			log("Unmasked Link URL: " + URL) 
		else:
			log("No Link found")
		return URL
	except:
		log("Nav URL error") 
		return URL
		
def getLinkName():
	""" 
	Get the current Link Name from the navigator objekt 
	"""
	obj = api.getNavigatorObject()
	linkName = ""
	try:
		if obj.role == 19: #link 
			linkName = obj.name
			log("Original Link Name: " + linkName) 
			linkName = validFilename(linkName)
			log("Modifyed Link Name: " + linkName) 
		else:
			log("No Link Name found")
		return linkName 
	except:
		log("Nav URL error") 
		return URL
		
def getMultimediaURLExtension():
	""" 
	Get the extension from the current URL 
	"""
	URL = getLinkURL()
	Ext = ""
	if URL != "":
		i = URL.rfind(".") 
		if i != -1:
			log("found point")
			l = len(URL)
			Ext = URL[i: l].lower()
	return Ext 
	
def isValidMultimediaExtension(Ext):
	""" 
	checks whether it is a valid multimedia extension.
	""" 
	r = False 
	for entry in MultimediaExtensions:
		if entry == Ext:
			return True 
	return r 

def getWebSiteTitle():
	""" 
		Get current web site title (new) 
	"""
	obj = api.getForegroundObject()
	t = ""
	try:
		t = obj.name
		log("Title Original: " + t)
		t = t.replace(" - YouTube", "")
	except:
		log("title: URL not found") 
		t = "Unknown_Title"
	return t
	
def makePrintable(s):
	"""
		replace not printable charakters with blank 
	"""
	result = ""
	l      = len(s) 
	for i in range(l):
		c = s[i]
		if c.isprintable():
			result = result + s[i]
		else:
			result = result + " " 
	return result 
	
def validFilename(s):
	"""
		returns a valid file name 
	"""
	result = ""
	l      = len(s) 
	for i in range(l):
		c = s[i]
		if c in invalidCharactersForFilename:
			result = result + " "
		else:
			result = result + c
	return result 
	
def is32bitMachine():
	"""
		Checks whether it is a 32bit (x86) operating system.
	"""
	result = True 
	try:
		res = str(platform.machine())
		log("operating system: " + res)
		if res == "AMD64":
			result = False 
	except:
		pass
	return result

def checkWritePermissions(folderPath):
	from stat import filemode 
	"""
		Checks whether the specified folder is writable.
		:param folderPath: Path to the folder you want to check.
		:return: True, if the folder is writable, otherwise False.
	"""
	try:
		mode = os.stat(folderPath).st_mode 
		accessRights = filemode(mode)
		if accessRights == "drwxrwxrwx":
			return True 
		else: 
			return False 
	except OSError:
		return False
		
def convertToMP(mpFormat, YouTubePath, OtherPath):
	log("*** Start " + mpFormat + " convert on " + getDate() + " " + getTime() + " ***")
	log("YouTube Path: " + str(YouTubePath))
	log("Other Path  : " + str(OtherPath))
	log("AppName: "  + getCurrentAppName()) 
	if not isBrowser():
		return 
	URL                = getCurrentDocumentURL() 
	validYouTubeURL    = True
	validMultimediaExt = False 
	if not URL:
		# Translators: No URL found in browser document, function terminate 
		ui.message(_("Document URL not found - exit"))
		return
	elif URL.find(".youtube.") == -1:
		# Translators: invalid youtube	 URL
		#ui.message(_("Invalid YouTube URL"))
		log("invalid YouTube URL") 
		validYouTubeURL = False 
	if validYouTubeURL:
		Title = getWebSiteTitle() 
		log(Title) 
		# Translators: save YouTube Video as MP3 or MP4
		ui.message(_("Save YouTube Video as {MultimediaFormat}").format(MultimediaFormat=mpFormat))
		if mpFormat == "mp3":
			cmd = [YouTubeEXE, "-x", "--audio-format", str(mpFormat), "--audio-quality", "0", "-o", "%(title)s.%(ext)s", URL]
		else: # if mp4 
			cmd = [YouTubeEXE, "-f", str(mpFormat), "-o", "%(title)s.%(ext)s", URL]
		if getINI("YouTubeDescription"): cmd.insert(1, "--write-description")
		ytThread = converterThread(cmd, YouTubePath)
		Wait = WaitThread(ytThread)
		Wait.start() 
	else: # Search for multimedia link 
		ext = getMultimediaURLExtension()
		ext = ext.replace(".", "")
		if ext == "":
			validMultimediaExt = False 
			log("No link extension found")
		else:
			validMultimediaExt = True 
			log("link extension found: " + ext)
			if not isValidMultimediaExtension(ext):
				validMultimediaExt = False 
				log("Extension is not a valid multimedia extension")
			else: 
				multimediaLinkURL  = getLinkURL()
				multimediaLinkName = getLinkName() + "." + str(mpFormat)
				cmd = [ConverterEXE, "-i", multimediaLinkURL, "-y", multimediaLinkName]
				log("cmd: " + str(cmd))
				# Translators: save multimedia link  as MP3 or MP4
				ui.message(_("Save link as {MultimediaFormat}").format(MultimediaFormat=mpFormat))
				multimediaThread = converterThread(cmd, OtherPath)
				Wait = WaitThread(multimediaThread)
				Wait.start() 
	if not validMultimediaExt and not validYouTubeURL:
		log("No YouTube Videoand no Multimedia Link found")
		# Translators: No YouTube video and no multimedia link found
		ui.message(_("No YouTube video and no multimedia link found"))

class AddOnPanel(SettingsPanel):
	title = AddOnSummary

	def makeSettings(self, sizer):
		helper = guiHelper.BoxSizerHelper(self, sizer=sizer)
		
		# this groupBox code is based on "NVDA Dev & Test Toolbox" from Cyrille Bougot, thanks 
		# Translators: This is a label for an edit field in the AVC Settings panel.
		resultFolderLabel = _("Result folder")
		groupSizer = wx.StaticBoxSizer(wx.VERTICAL, self, label=resultFolderLabel)
		groupBox = groupSizer.GetStaticBox()
		groupHelper = helper.addItem(gui.guiHelper.BoxSizerHelper(self, sizer=groupSizer))
		# Translators: The label of a button to browse for a directory
		browseText = _("&Browse...")
		# Translators: The title of the dialog presented when browsing for the directory.
		dirDialogTitle = _("Select a directory")
		directoryPathHelper = gui.guiHelper.PathSelectionHelper(groupBox, browseText, dirDialogTitle)
		directoryEntryControl = groupHelper.addItem(directoryPathHelper)
		self.resultFolderEdit       = directoryEntryControl.pathControl
		self.resultFolderEdit.Value = getINI("ResultFolder")

		# Translators: Checkbox name for sound signal feedback in the configuration dialog
		self.BeepConvertingChk = helper.addItem(wx.CheckBox(self, label=_("&beep while converting")))
		self.BeepConvertingChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.BeepConvertingChk.Value = getINI("BeepWhileConverting")
		# Translators: Checkbox name in the configuration dialog
		self.YouTubeDescriptionChk = helper.addItem(wx.CheckBox(self, label=_("&Write YoutTube description file")))
		self.YouTubeDescriptionChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.YouTubeDescriptionChk.Value = getINI("YouTubeDescription")
		# Translators: Checkbox name in the configuration dialog
		self.LoggingChk = helper.addItem(wx.CheckBox(self, label=_("Write &report file")))
		self.LoggingChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.LoggingChk.Value = getINI("Logging")

	def onSave(self):
		folder = self.resultFolderEdit.Value
		folder = folder.strip() 
		if folder.endswith("\\"):
			folder = folder[0:len(folder)-1]
		if not os.path.exists(folder) or folder == "":
			# Translator: Title of error dialog "Folder does not exist"
			errorNotExistsDialogTitle = _("Error - Folder does not exist")
			# Translator: content  of error dialog "Folder does not exist"
			errorNotExistsDialogText  = _(
				"The folder ({folderPath}) does not exist.\n" 
				"Please check the path or use the Browse button.\n"
				"The changes were not saved."
			).format(folderPath=folder)
			gui.messageBox(errorNotExistsDialogText, errorNotExistsDialogTitle, wx.OK | wx.ICON_ERROR)
		else:
			if checkWritePermissions(folder):
				setINI("BeepWhileConverting", bool(self.BeepConvertingChk.Value))
				setINI("YouTubeDescription",  bool(self.YouTubeDescriptionChk.Value))
				setINI("Logging",             bool(self.LoggingChk.Value))
				setINI("ResultFolder",        folder)
				CheckFolders()
			else:
				octCode = str(oct(os.stat(folder).st_mode))
				# Translator: Title of error dialog "No write access to folder"
				errorNotAccessDialogTitle = _("Error - No write access to folder")
				# Translator: content  of error dialog "No write access to folder"
				errorNotAccessDialogText  = _(
					"No write access to folder {folderPath} (Code: {code})\n" 
					"Please select a folder with write access.\n"
					"The changes were not saved."
			)	.format(folderPath=folder, code=octCode)
				gui.messageBox(errorNotAccessDialogText, errorNotAccessDialogTitle, wx.OK | wx.ICON_ERROR)

	def onChk(self, event):
		pass 

class converterThread(threading.Thread):

	def __init__(self, cmd, Path):
		super().__init__()
		self.cmd = cmd 
		self.Path = Path 
		self.stopSignal = False 

	def run(self):
		global processID
		log(" ".join(self.cmd))
		log("Start: " + getTime())  
		StartTime = int(time.time())
		# Si - Process should run in the background
		si = subprocess.STARTUPINFO()
		si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
		#p = subprocess.Popen(self.cmd, cwd=self.Path, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, startupinfo=si, encoding="unicode_escape", text=True)
		self.p = subprocess.Popen(self.cmd, cwd=self.Path, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=si, encoding="unicode_escape")
		processID = self.p.pid
		log("Process ID: " + str(processID))
		stdout, stderr = self.p.communicate()
		log("StandardError: " + str(stderr))
		log("ReturnCode: " + str(self.p.returncode))
		log("stdOut:\n" + str(stdout))
		self.p.wait() 
		EndTime = int(time.time())
		Sec = EndTime - StartTime
		log("End  : " + getTime() + " (" + str(Sec) + " Sec)")

	def terminateProcess():
		global processID
		log("Terminate process by user")
		if processID:
			# Translators: The converting process will terminate by user
			ui.message(_("The converting process will terminate."))
			try:
				os.kill(processID, 4)
				processID = None 
				log("terminate process successful")
			except:
				log("Error: terminating process")
		else:
			log("No converting task active.")
			# Translators: Canceling process by user. No converting task active.
			ui.message(_("No converting task active."))

class WaitThread(threading.Thread):

	def __init__(self, myThread):
		super().__init__()
		self.myThread = myThread

	def run(self):
		self.myThread.start() 
		i = 0
		while self.myThread.is_alive():
			time.sleep(0.1)
			i += 1
			if i == 10:
				#tones.beep(500, 100)
				PlayWave("buisy")
				i = 0 
		self.myThread.join()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = AddOnSummary
	
	def __init__(self):
		super(globalPluginHandler.GlobalPlugin, self).__init__()
		if getINI("ResultFolder") == "":
			setINI("ResultFolder", DownloadPath) 
		CheckFolders()
		# Add a section in NVDA configurations panel
		NVDASettingsDialog.categoryClasses.append(AddOnPanel)

	def terminate(self):
		try:
			NVDASettingsDialog.categoryClasses.remove(AddOnPanel)
		except Exception:
			pass

	@script(
		description=_("Save as MP3"),
		gesture="kb:NVDA+y",
		speakOnDemand=True
	)
	def script_ConvertToMP3(self, gesture):
		convertToMP("mp3", YouTubeAudio, OtherAudio)

	@script(
		description=_("Save as MP4"),
		gesture="kb:NVDA+shift+y",
		speakOnDemand=True
	)
	def script_YouTubeToMP4(self, gesture):
		convertToMP("mp4", YouTubeVideo, OtherVideo)

	@script(
		description=_("Open result folder"),
		gesture="kb:NVDA+control+y",
		speakOnDemand=True
	)
	def script_ResultFolder(self, gesture):
		DownloadPath = getINI("ResultFolder")
		log("open the result folder: " + DownloadPath)
		os.startfile(DownloadPath)

"""
	@script(
		description="Only for test",
		gesture="kb:NVDA+l"
	)
	def script_test(self, gesture):
		pass
"""