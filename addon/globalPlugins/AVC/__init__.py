# -*- coding: UTF-8 -*-

"""
 Audio Video Converter for NVDA 
 This file is covered by the GNU General Public License.
 See the file COPYING for more details.
 Copyright (C) 2022/2025 Rainer Brell nvda@brell.net 

 *** 22. Nov 2022, Version 1.0
 NVDA+y         Converts to MP3 
 NVDA+shift+y   Converts to MP4
 nvda+alt+y     Saves the subtitle of a YouTube video
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
  *** 2025.06.16:
  - Look for a new yt-dlp.exe once every day
  - Ready vor NVDA 2025.1 
  - more log informations 
  *** 2025.06.30:
  - nvda+alt+y, Saves the subtitle of a YouTube video
 
 *** In Progress for future  *** 
 - The conversion process can be killed 
 - Function for conversion into  windows explorer 
 - Password-protected access to YouTube videos
 - Audio export from the video container
  """

import globalPluginHandler
from scriptHandler import script
import ui
import urllib 
import signal 
import globalVars
import gui
from gui.settingsDialogs import NVDASettingsDialog, SettingsPanel
from gui import guiHelper
from .skipTranslation import translate
import wx 
import tones 
import config 
import datetime 
import time 
import api
import os 
import core 
import nvwave 
import controlTypes
import languageHandler
import scriptHandler
import subprocess
import threading
import platform
import addonHandler

addonHandler.initTranslation()

AddOnSummary     = addonHandler.getCodeAddon().manifest['summary']
AddOnName        = addonHandler.getCodeAddon().manifest['name']
AddOnVersion     = addonHandler.getCodeAddon().manifest['version']
AddOnPath        = os.path.dirname(__file__)
ToolsPath        = AddOnPath + "\\Tools\\"
SoundPath        = AddOnPath + "\\sounds\\"
AppData          = os.environ["APPDATA"]
DownloadPath     = AppData + "\\AVC-Results"
LogFile          = DownloadPath + "\\log.txt"
LastUpdateFile   = DownloadPath + "\\LastUpdate.txt"
YouTubeVideo     = DownloadPath + "\\YouTubeVideo\\"
YouTubeAudio     = DownloadPath + "\\YouTubeAudio\\"
OtherVideo       = DownloadPath + "\\Video\\"
OtherAudio       = DownloadPath + "\\Audio\\"
SubtitleFolder   = DownloadPath + "\\Subtitle\\"
YouTubeEXE       = AddOnPath + "\\Tools\\yt-dlp.exe"
ConverterEXE     = AddOnPath + "\\Tools\\ffmpeg.exe"
sectionName      = AddOnName 
processID        = None 
DelayAfterStart  = 120 # secounds 
ShortLanguage = languageHandler.getLanguage()[:2]

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
	"\"",
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
		"YouTubePlaylist":              "boolean(default=False)",
		"YouTubeChannel":               "boolean(default=False)",
		"YouTubeSubtitle":              "boolean(default=False)",
		"SubtitleLanguage":             f"string(default='{ShortLanguage}')",
		"Logging":                      "boolean(default=False)"
	}
	config.conf.spec[sectionName] = confspec

initConfiguration()

def should_update():
	"""
		Check whether you want to update yt-dlp.exe.
		A new update should be checked once a day.
	"""
	# Load the date of the last program start from a file
	try:
		with open(LastUpdateFile, "r") as file:
			last_run_date = file.read().strip()
	except FileNotFoundError:
		last_run_date = None
	# get Current date
	current_date = datetime.datetime.now().date()
	log(f"Current date: {current_date}, last date: {last_run_date}")
	# Check whether the update was already carried out today
	if last_run_date != str(current_date):
		doUpdate = True 
	else: 
		doUpdate = False 
	# Save the current date to the file
	with open(LastUpdateFile, "w") as file:
		file.write(str(current_date))
	return doUpdate

def is_online():
	"""
		Check if the system is online
	"""
	import socket
	try:
		# Try to connect to a known host (e.g. Google DNS)
		socket.create_connection(("8.8.8.8", 53))
		return True 
	except OSError:
		return False


def update_YouTubeEXE():
	"""
		Runs an update of yt-dlp.exe in the background
	"""
	log("Try to update the yt-dlp.exe")
	try:
		cmd = YouTubeEXE
		result = subprocess.run([cmd, "-U"], capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
		output = result.stdout
		log(output)
		return output
	except Exception as e:
		log(str(e))
		return str(e)
		
def convert_and_show_subtitle():
	msg, title = extract_subtitles_as_text(find_latest_srv1_file(SubtitleFolder))
	#ui.message(msg) 
	#ui.browseableMessage(msg, title=title, isHtml=False)

def save_subtitle(cmd):
	"""
		Saves the subtitle in the background
	"""
	log("Save the subtitel")
	log("Command: " + " ".join(cmd))
	try:
		result = subprocess.run(
			cmd, 
			capture_output=True, 
			text=True, 
			creationflags=subprocess.CREATE_NO_WINDOW
		)
		output = result.stdout
		log("Returncode: " + str(self.p.returncode))
		log("Returncode: " + str(result.returncode) )
		log("Error     : " + str(result.stderr))
		log("Output    : " + output)
		timer_thread = threading.Timer(1, convert_and_show_subtitle)
		timer_thread.start()
	except Exception as e:
		log(str(e))
		
def find_latest_srv1_file(directory_path):
	"""
	Find the most recently modified file with the .SRV1 extension
	in a given directory. Returns the file path or None if none found.
	"""
	from pathlib import Path
	path = Path(directory_path)
	srv1_files = list(path.glob("*.SRV1"))
	if not srv1_files:
		return None  # No .SRV1 files found
	# Sort by last modified time (mtime)
	latest_file = max(srv1_files, key=lambda f: f.stat().st_mtime)
	return latest_file
	
def extract_subtitles_as_text(srv1_file_path):
	"""
	Converts a .srv1 subtitle file to plain text and saves it as .txt
	with the same base filename in the same directory.
	:param srv1_file_path: Path to the input .srv1 file
	returns the text content and the title of the file
	Deletes the SRV1 file
	"""
	from xml.etree import ElementTree as ET
	if not srv1_file_path: 
			return 
	try:
		tree = ET.parse(srv1_file_path)
		root = tree.getroot()
		lines = [node.text for node in root.iter("text") if node.text]
		txt_file_path = os.path.splitext(srv1_file_path)[0] + ".txt"
		with open(txt_file_path, "w", encoding="utf-8") as out:
			out.write("\n".join(lines))
		log(f"Transcript saved under: {txt_file_path}")
		filename = os.path.splitext(os.path.basename(txt_file_path))[0]
		os.remove(srv1_file_path)
		return "\n".join(lines), filename
	except ET.ParseError as e:
		log(f"Error when parsing the XML file: {e}")
		return None 

def YouTubeExe_update():
	"""
		Updated to the latest version of the YouTube downloader
	""" 
	if is_online():
		thread = threading.Thread(target=update_YouTubeEXE)
		thread.start()
		thread.join()
		
def CheckYouTubeEXE():
	"""
		If the system is online, it will look for a new yt-dlp.exe when it is first started
	"""
	if is_online(): 
			if should_update():
				log(f"Search for a new yt-dlp.exe after {DelayAfterStart} seconds") 
				timer_thread = threading.Timer(DelayAfterStart, YouTubeExe_update)
				timer_thread.start()
				
def getActiveProfile ():
	activeProfile = config.conf.profiles[-1].name
	if not activeProfile:
		# Message translated in NVDA core.
		activeProfile = translate("normal configuration")
	return activeProfile

def getINI(key):
	"""  get nvda.ini value """ 
	value = config.conf[sectionName][key]
	log(f"Get {value} from key {key} of section {sectionName}")
	return value 

def setINI(key, value):
	"""  set nvda.ini value """ 
	try:
		config.conf[sectionName][key] = value
		log(f"set {value} to   key {key} of section {sectionName}")
	except:
		log("Error when writing the ini")

def log(s):
	""" 
		Write to log file 
	""" 
	if config.conf[sectionName]["Logging"]:
		CurrentTime = datetime.datetime.now().time()
		CurrentTime = CurrentTime.strftime("%H:%M:%S")
		try:
			s = makePrintable(s)
			s = CurrentTime + ": " + s 
			lines = s.splitlines()
			log = open(LogFile, "a")
			for line in lines:
				log.write(str(line))
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
	global LastUpdateFile
	global YouTubeVideo
	global YouTubeAudio
	global OtherVideo
	global OtherAudio
	global SubtitleFolder
	# set new global vars 
	DownloadPath     = getINI("ResultFolder")
	LogFile          = DownloadPath + "\\log.txt"
	LastUpdateFile   = DownloadPath + "\\LastUpdate.txt"
	YouTubeVideo     = DownloadPath + "\\YouTubeVideo\\"
	YouTubeAudio     = DownloadPath + "\\YouTubeAudio\\"
	OtherVideo       = DownloadPath + "\\Video\\"
	OtherAudio       = DownloadPath + "\\Audio\\"
	SubtitleFolder   = DownloadPath + "\\Subtitle\\"
	# create folders 
	createFolder(DownloadPath)
	createFolder(YouTubeVideo)
	createFolder(YouTubeAudio)
	createFolder(OtherAudio)
	createFolder(OtherVideo)
	createFolder(SubtitleFolder)

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
	for c in s:
		if c in invalidCharactersForFilename:
			result = result + ""
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
	tempfile  = os.path.join(folderPath, 	"HierKommtDieMausUndDuBistAus.HugoUndErna")
	"""
		Checks whether the specified folder is writable.
		This is not a good solution, but more reliable than that: os.access or os.stat
		:param folderPath: Path to the folder you want to check.
		:return: True, if the folder is writable, otherwise False.
	"""
	try:
		with open(tempfile, "w") as f:
			f.write("Hurra")
		os.remove(tempfile)
		return True 
	except IOError:
		return False
		
def checkWritePermissions_old(folderPath):
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
		if getINI("YouTubeDescription"): cmd.insert(-1, "--write-description")
		if getINI("YouTubeSubtitle"): 
				cmd.insert(-1, "--write-auto-sub")
				cmd.insert(-1, "--sub-langs")
				cmd.insert(-1, "DE")
				cmd.insert(-1, "--sub-format")
				cmd.insert(-1, "srv1")
		if not getINI("YouTubeChannel"): 
				cmd.insert(-1, "--max-downloads")
				cmd.insert(-1, "1")
		if not getINI("YouTubePlaylist"): cmd.insert(-1, "--no-playlist")
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
				log("cmd: " + str(" ".join(cmd)))
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
		
		# Translators: Checkbox name in the configuration dialog  for video description 
		self.YouTubeDescriptionChk = helper.addItem(wx.CheckBox(self, label=_("&Write YoutTube description file")))
		self.YouTubeDescriptionChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.YouTubeDescriptionChk.Value = getINI("YouTubeDescription")
		
		# Translators: Checkbox name in the configuration dialog  for youtube subtitle 
		#self.YouTubeSubtitleChk = helper.addItem(wx.CheckBox(self, label=_("Always generate youtube &subtitle (needs more time)")))
		#self.YouTubeSubtitleChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		#self.YouTubeSubtitleChk.Value = getINI("YouTubeSubtitle")
		
		# Translators: Checkbox name in the configuration dialog  for youtube Playlist 
		self.YouTubePlaylistChk = helper.addItem(wx.CheckBox(self, label=_("Download the &play list completely")))
		self.YouTubePlaylistChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.YouTubePlaylistChk.Value = getINI("YouTubePlaylist")
		
		# Translators: Checkbox name in the configuration dialog  for youtube Channel 
		self.YouTubeChannelChk = helper.addItem(wx.CheckBox(self, label=_("Download the &channel completely")))
		self.YouTubeChannelChk.Bind(wx.EVT_CHECKBOX, self.onChk)
		self.YouTubeChannelChk.Value = getINI("YouTubeChannel")

 		# Translators: Checkbox name in the configuration dialog for AVC logging 
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
				#setINI("YouTubeSubtitle",  bool(self.YouTubeSubtitleChk.Value))
				setINI("YouTubePlaylist",  bool(self.YouTubePlaylistChk.Value))
				setINI("YouTubeChannel",  bool(self.YouTubeChannelChk.Value))
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

class SubtitleThread(threading.Thread):

	def __init__(self, cmd, path):
		super().__init__()
		self.cmd  = cmd 
		self.Path = path 

	def run(self):
		log("Save the subtitel")
		log("Command: " + " ".join(self.cmd))
		# Si - Process should run in the background
		si = subprocess.STARTUPINFO()
		si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
		self.p = subprocess.Popen(
			self.cmd, 
			cwd=self.Path, 
			stdin=subprocess.PIPE, 
			stdout=subprocess.PIPE, 
			stderr=subprocess.PIPE, 
			startupinfo=si, 
			encoding="unicode_escape"
		)
		output = self.p.stdout
		log("Returncode: " + str(self.p.returncode))
		log("Error     : " + str(self.p.stderr))
		log("Output    : " + str(output))
		log("**** neu ***")
		# Lies stdout und stderr komplett
		stdout_data, stderr_data = self.p.communicate()
		log("Returncode: " + str(self.p.returncode))
		log("Error     :")
		for line in stderr_data.splitlines():
			log(line)
		og("Output    :")
		for line in stdout_data.splitlines():
			log(line)
		self.p.wait() 
		msg, title = extract_subtitles_as_text(find_latest_srv1_file(SubtitleFolder))

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
		# if globalVars.appArgs.secure: return 
		if getINI("ResultFolder") == "":
			setINI("ResultFolder", DownloadPath) 
		CheckFolders()
		log("AVC Version: " + AddOnVersion)
		CheckYouTubeEXE()
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
		
	@script(
		description=_("Saves the subtitle of a YouTube video"),
		gesture="kb:NVDA+alt+y",
		speakOnDemand=True
	)
	def script_SaveSubTitle(self, gesture):
		CurrentLanguage   = getINI("SubtitleLanguage")
		LocalizedLanguage = languageHandler.getLanguageDescription(CurrentLanguage)
		log("Current language: " + LocalizedLanguage)
		if not isBrowser():
				return 
		URL                = getCurrentDocumentURL()
		validYouTubeURL    = True
		if not URL:
			# Translators: No URL found in browser document, function terminate 
			ui.message(_("Document URL not found - exit"))
			return
		elif URL.find(".youtube.") == -1:
			log("invalid YouTube URL") 
			validYouTubeURL = False 
		if validYouTubeURL:
			Title = getWebSiteTitle() 
			log(Title) 
			# Translators: Save the subtitle in the specified language
			ui.message(_("Save subtitle in {language}").format(language=LocalizedLanguage))
			cmd = [
				YouTubeEXE, 
				"--write-auto-sub", 
				"--sub-langs", 
				CurrentLanguage, 
				"--skip-download", 
				"--sub-format", "srv1", 
				"--paths", SubtitleFolder, 
				"--output", "%(title)s.%(ext)s",
				URL
			]
			SubtThread = SubtitleThread(cmd, SubtitleFolder)
			Wait = WaitThread(SubtThread)
			Wait.start() 
			#SubtitleThread = threading.Thread(target=save_subtitle, args=(cmd,))
			#SubtitleThread.start()
			#SubtitleThread.join()
		else:
			log("No YouTube Video Link found")
			# Translators: No YouTube video link found
			ui.message(_("No YouTube video link found"))

	@script(
		description="Only for test",
		gesture="kb:NVDA+l"
	)
	def script_test(self, gesture):
		msg = "Hugo"
		ui.message(msg) 
		