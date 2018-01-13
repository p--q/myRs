#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.document import XDocumentEventListener
from myrs import consts  # 相対インポートは不可。
def myRs(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。この関数ではXSCRIPTCONTEXTは使えない。  
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
# 	doc.addChangesListener(ChangesListener(doc))  # ChangesListener	
# 	doc.addDocumentEventListener(DocumentEventListener(doc, tdocimport, modulefolderpath))  # DocumentEventListener	
	controller = doc.getCurrentController()  # コントローラの取得。
# 	selectionchangelistener = SelectionChangeListener(controller)  # SelectionChangeListener
# 	controller.addActivationEventListener(ActivationEventListener(controller))  # ActivationEventListener
# 	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(controller, selectionchangelistener))  # EnhancedMouseClickHandler
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, smgr, doc))  # コントローラにContextMenuInterceptorを登録する。右クリックの時の対応。
# class DocumentEventListener(unohelper.Base, XDocumentEventListener):
# 	def __init__(self, subj, tdocimport, modulefolderpath):
# 		self.subj = subj
# 		self.args = tdocimport, modulefolderpath
# 	def documentEventOccured(self, documentevent):
# 		tdocimport, modulefolderpath = self.args
# 		eventname = documentevent.EventName
# 		if eventname=="OnUnload":  # ドキュメントを閉じる時。
# 			tdocimport.remove_meta(modulefolderpath)  # modulefolderpathをメタパスから除去する。
# 	def disposing(self, eventobject):	
# 		self.subj.removeDocumentEventListener(self)
# class ActivationEventListener(unohelper.Base, XActivationEventListener):
# 	def __init__(self, controller):
# 		self.controller = controller
# 	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
# 		sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
# 		sheetname = sheet.getName()  # アクティブシート名を取得。
# 		if sheetname=="一覧":
# 			ichiran.activeSpreadsheetChanged(sheet)
# 	def disposing(self, eventobject):
# 		self.controller.removeActivationEventListener(self)	
# class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
# 	def __init__(self, controller, colors, modules, selectionchangelistener):
# 		self.controller = controller
# 		self.args = colors, modules, selectionchangelistener
# 	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。
# 		colors, modules, selectionchangelistener = self.args
# 		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
# 		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
# 			sheet = target.getSpreadsheet()
# 			sheetname = sheet.getName()  # アクティブシート名を取得。
# 			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時
# 				if sheetname=="一覧":
# 					ichiran.singleClick(colors, self.controller, target)
# 					
# 				# ここで罫線を引く
# 				
# # 				controller.addSelectionChangeListener(self.selectionchangelistener)
# 			
# 			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 				celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
# # 				if controller.hasFrozenPanes():  # 表示→セルの固定、がされている時。
# # 					splitrow = controller.getSplitRow()
# # 					splitcolumn = controller.getSplitColumn()
# 				return False  # セル編集モードにしない。
# 		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
# 	def mouseReleased(self, enhancedmouseevent):
# 		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
# 			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時		
# 				try:
# 					self.controller.removeSelectionChangeListener(self.selectionchangelistener)		
# 				except:
# 					pass
# 		return True
# 	def disposing(self, eventobject):
# 		self.controller.removeEnhancedMouseClickHandler(self)	
# class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
# 	def __init__(self, modules, controller):
# 		self.controller = controller
# 	def selectionChanged(self, eventobject):
# 		controller = self.controller
# 		selection = controller.getSelection()  # 選択しているオブジェクトを取得。
# 	def disposing(self, eventobject):
# 		self.controller.removeSelectionChangeListener(self)		
# class ChangesListener(unohelper.Base, XChangesListener):
# 	def __init__(self, modules, doc):
# 		self.subj = doc
# 	def changesOccurred(self, changesevent):
# 		changes = changesevent.Changes
# 		for change in changes:
# 			accessor = change.Accessor
# 			if accessor=="cell-change":  # セルの内容が変化した時。
# 				cell = change.ReplacedElement  # 変化したセルを取得。				
# 	def disposing(self, eventobject):
# 		self.doc.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, smgr, doc):
		self.args = getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True) 
		baseurl = self.args 
		controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
		target = controller.getSelection()  # 選択しているオブジェクトを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
		name = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
		addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
		if name=="cell":  # セルのとき
			del contextmenu[:]  # contextmenu.clear()は不可。
			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
				addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format(toBlue.__name__)})  # 引数のない関数名を渡す。
			elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
				addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format(toRed.__name__)})  # 引数のない関数名を渡す。
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		elif name=="rowheader":  # 行ヘッダーのとき。
			del contextmenu[:]  # contextmenu.clear()は不可。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
		elif name=="colheader":  # 列ヘッダーの時。
			pass  # contextmenuを操作しないとすべての項目が表示されない。
		elif name=="sheettab":  # シートタブの時。
			del contextmenu[:]  # contextmenu.clear()は不可。
			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
		return EXECUTE_MODIFIED	  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def toBlue():
	colors = consts.COLORS
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	target = doc.getCurrentSelection()
	target.setPropertyValue("CellBackColor", colors["ao"])  # 背景を青色にする。
def toRed():
	colors = consts.COLORS
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	target = doc.getCurrentSelection()
	target.setPropertyValue("CellBackColor", colors["aka"])  # 背景を青色にする。


def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	i = 0  # インデックスを初期化する。
	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
		nonlocal i
		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
		i += 1  # インデックスを増やす。
	return addMenuentry
def getBaseURL(ctx, smgr, doc):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	filepath = modulepath.replace(ucp, "")  #  ucpを除去。
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
	contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
	macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
	location = "document"  # マクロの場所。	
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	