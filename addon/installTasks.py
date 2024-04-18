# -*- coding: UTF-8 -*-
"""
	installTasks.py for Audio Video Converter
	Copyright 2024 Rainer Brell	, released under gPL.
	This file is covered by the GNU General Public License.
"""

def decompressZipFile(zipFile, zipPath):
	"""
		Extracts all files from a ZIP file to the specified destination directory.
		:param zipFile:: Path to the ZIP file
		:param zipPath : Directory to which the files should be extracted
	"""
	import zipfile 
	try: 
		with zipfile.ZipFile(zipFile, 'r') as zip_ref:
			zip_ref.extractall(zipPath)
		return True 
	except: 
		return False 

def myStatistic():
	"""
		I would like to use these statistics to find out which countries install 
		my add-on in which version in order to possibly offer a translation in these languages.
	""" 
	import urllib 
	import os 
	import addonHandler
	import languageHandler 
	import datetime 
	from versionInfo import version as nvdaVersion 
	addonDir     = os.path.dirname(__file__)
	addonName    = addonHandler.Addon(addonDir).manifest["name"]
	addonVersion = addonHandler.Addon(addonDir).manifest["version"]
	lang         = languageHandler.getLanguage().lower()
	fileName     = (addonName + addonVersion + ".csv")
	date         = datetime.datetime.now().strftime("%Y.%m.%d")
	time         = datetime.datetime.now().strftime("%H:%M:%S")
	line         = lang + ";" + nvdaVersion + ";" + date + ";" + time 
	base_url     = "https://nvda.brell.net/statistic"
	params       = {
		"param1": fileName,
		"param2": line
	}
	url = f"{base_url}?fileName={params['param1']}&line={params['param2']}"
	try: 
		urllib.request.urlopen(url)
	except:
		pass

def onInstall():
	myStatistic()
	import os 
	import gui 
	import wx 
	import gettext
	_ = gettext.gettext
	zipPath = os.path.dirname(__file__) + "\\globalPlugins\\AVC\\Tools\\"
	zipFile = zipPath + "ffmpeg.zip"
	if not decompressZipFile(zipFile, zipPath):
		# Translators: Title of the window if an error occurred during unpacking.
		title: str = _("Attention - error when unpacking")
		# Translators: Contents of the window in the event of an error during unpacking
		msg: str = _(
			"The file {file} could not be unpacked correctly in the folder {folder}. \n"
			"Try it yourself locally or contact the author of this extension."
		).format(
			file=zipFile, 
			folder=zipPath
		)
		gui.messageBox(msg, title, wx.OK | wx.ICON_ERROR)
