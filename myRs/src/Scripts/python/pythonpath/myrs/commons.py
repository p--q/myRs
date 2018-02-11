#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.datatransfer import XTransferable
from com.sun.star.datatransfer import DataFlavor  # Struct
from com.sun.star.datatransfer import UnsupportedFlavorException
COLORS = {\
# 		"lime": 0x00FF00,\
		"magenta3": 0xFF00FF,\
		"black": 0x000000,\
# 		"blue3": 0x0000FF,\
		"skyblue": 0x00CCFF,\
		"silver": 0xC0C0C0,\
# 		"red3": 0xFF0000,\
		"violet": 0x9999FF,\
		"cyan10": 0xCCFFFF}  # 色の16進数。	
class TextTransferable(unohelper.Base, XTransferable):
	def __init__(self, txt):  # クリップボードに渡す文字列を受け取る。
		self.txt = txt
		self.unicode_content_type = "text/plain;charset=utf-16"
	def getTransferData(self, flavor):
		if flavor.MimeType.lower()!=self.unicode_content_type:
			raise UnsupportedFlavorException()
		return self.txt
	def getTransferDataFlavors(self):
		return DataFlavor(MimeType=self.unicode_content_type, HumanPresentableName="Unicode Text"),  # DataTypeの設定方法は不明。
	def isDataFlavorSupported(self, flavor):
		return flavor.MimeType.lower()==self.unicode_content_type
