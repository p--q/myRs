#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# embeddedmacro.pyから呼び出した関数ではXSCRIPTCONTEXTは使えない。デコレーターも使えない。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)でブレークする。
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.document import XDocumentEventListener
from com.sun.star.table import BorderLine2  # Struct
from com.sun.star.table import BorderLineStyle  # 定数
from myrs import commons, ichiran, karute, keika, rireki, taiin, yotei  # 相対インポートは不可。

def myRs(tdocimport, modulefolderpath, xscriptcontext):  # 引数は文書のイベント駆動用。この関数ではXSCRIPTCONTEXTは使えない。  
	
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	
	
	
	ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	doc.addChangesListener(ChangesListener())  # ChangesListener	
	doc.addDocumentEventListener(DocumentEventListener(tdocimport, modulefolderpath))  # DocumentEventListener	
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.addSelectionChangeListener(SelectionChangeListener())
	controller.addActivationEventListener(ActivationEventListener())  # ActivationEventListener
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(controller))  # EnhancedMouseClickHandler。このリスナーのメソッドの引数からコントローラーを取得する方法がない。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, smgr, doc))  # コントローラにContextMenuInterceptorを登録する。右クリックの時の対応。
class DocumentEventListener(unohelper.Base, XDocumentEventListener):
	def __init__(self, tdocimport, modulefolderpath):
		self.args = tdocimport, modulefolderpath
	def documentEventOccured(self, documentevent):
		tdocimport, modulefolderpath = self.args
		eventname = documentevent.EventName
		if eventname=="OnUnload":  # ドキュメントを閉じる時。
			tdocimport.remove_meta(modulefolderpath)  # modulefolderpathをメタパスから除去する。
	def disposing(self, eventobject):	
		eventobject.Source.removeDocumentEventListener(self)
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
		sheetname = sheet.getName()  # アクティブシート名を取得。
		if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
			pass
		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
			pass
		elif sheetname=="一覧":
			ichiran.activeSpreadsheetChanged(sheet)
		elif sheetname=="予定":
			pass
		elif sheetname=="退院":
			pass
		elif sheetname=="履歴":
			pass
	def disposing(self, eventobject):
		eventobject.Source.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, controller):
		self.controller = controller
		colors = commons.COLORS
		noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
		firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["clearblue"])
		secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["magenta"])
		self.args = noneline, firstline, secondline
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
			sheet = target.getSpreadsheet()
			sheetname = sheet.getName()
			if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
				return True
			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
				return True
			elif sheetname=="一覧":
				return ichiran.mousePressed(enhancedmouseevent, sheet, target, self.args)
			elif sheetname=="予定":
				return True
			elif sheetname=="退院":
				return True
			elif sheetname=="履歴":
				return True
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。マウスボタンを離した時。複数セルを選択した後でもtargetはセルしか入らない。
		if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # targetがチャートの時がありうる?
			sheet = target.getSpreadsheet()
			sheetname = sheet.getName()
			if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
				return True
			elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
				return True
			elif sheetname=="一覧":
				pass
# 				doc = self.controller.getModel()  # ドキュメントを取得。モデルを渡すと選択セルの変更が反映されていない可能性がある。
# 				return ichiran.mouseReleased(enhancedmouseevent, doc, sheet, target, self.args)
			elif sheetname=="予定":
				return True
			elif sheetname=="退院":
				return True
			elif sheetname=="履歴":
				return True
		return True
	def disposing(self, eventobject):  # eventobject.SourceはNone。
		self.controller.removeEnhancedMouseClickHandler(self)	
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self):
		colors = commons.COLORS
		noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
		firstline = BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["clearblue"])
		secondline =  BorderLine2(LineStyle=BorderLineStyle.DASHED, LineWidth=62, Color=colors["magenta"])
		self.args = noneline, firstline, secondline	
	def selectionChanged(self, eventobject):  # マウスから呼び出した時の反応が遅い。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		controller = eventobject.Source
		sheet = controller.getActiveSheet()
		sheetname = sheet.getName()  # アクティブシート名を取得。		
		if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
			pass
		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
			pass
		elif sheetname=="一覧":
			ichiran.selectionChanged(controller, sheet, self.args)
		elif sheetname=="予定":
			pass
		elif sheetname=="退院":
			pass
		elif sheetname=="履歴":
			pass			
	def disposing(self, eventobject):
		eventobject.Source.removeSelectionChangeListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def changesOccurred(self, changesevent):  # Sourceにはドキュメントが入る。
		doc = changesevent.Source
		controller = doc.getCurrentController()
		sheet = controller.getActiveSheet()
		sheetname = sheet.getName()  # アクティブシート名を取得。
		if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
			pass
		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
			pass
		elif sheetname=="一覧":
			pass
		elif sheetname=="予定":
			pass
		elif sheetname=="退院":
			pass
		elif sheetname=="履歴":
			pass		
		
		
		
		
# 		changes = changesevent.Changes
# 		for change in changes:
# 			accessor = change.Accessor
# 			if accessor=="cell-change":  # セルの内容が変化した時。
# 				cell = change.ReplacedElement  # 変化したセルを取得。		
				
						
	def disposing(self, eventobject):
		eventobject.Source.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, smgr, doc):
		self.args = getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
# 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
		baseurl = self.args 
		controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
		contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
		addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
		sheet = controller.getActiveSheet()  # アクティブシートを取得。
		sheetname = sheet.getName()  # シート名を取得。
		if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
			karute.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
			keika.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		elif sheetname=="一覧":
			ichiran.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		elif sheetname=="予定":
			yotei.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		elif sheetname=="退院":
			taiin.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		elif sheetname=="履歴":
			rireki.notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname)
		return EXECUTE_MODIFIED	  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# ContextMenuInterceptorのnotifyContextMenuExecute()メソッドで設定したメニュー項目から呼び出される関数。関数名変更不可。動的生成も不可。
def entry1():
	invokeMenuEntry(1)
def entry2():
	invokeMenuEntry(2)	
def entry3():
	invokeMenuEntry(3)	
def entry4():
	invokeMenuEntry(4)
def entry5():
	invokeMenuEntry(5)
def entry6():
	invokeMenuEntry(6)
def entry7():
	invokeMenuEntry(7)
def entry8():
	invokeMenuEntry(8)
def entry9():
	invokeMenuEntry(9)	
	
	
def invokeMenuEntry(entrynum):  # コンテクストメニュー項目から呼び出された処理をシートごとに振り分ける。
	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
	selection = doc.getCurrentSelection()  # セル(セル範囲)またはセル範囲、セル範囲コレクションが入るはず。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲コレクション以外の時。
		sheet = selection.getSpreadsheet()  # シートを取得。
		sheetname = sheet.getName()  # シート名を取得。
		if sheetname.isdigit():  # シート名が数字のみの時カルテシート。
			karute.contextMenuEntries(selection, entrynum)
		elif sheetname.endswith("経"):  # シート名が「経」で終わる時は経過シート。
			keika.contextMenuEntries(selection, entrynum)
		elif sheetname=="一覧":
			ichiran.contextMenuEntries(selection, entrynum)
		elif sheetname=="予定":
			yotei.contextMenuEntries(selection, entrynum)
		elif sheetname=="退院":
			taiin.contextMenuEntries(selection, entrynum)
		elif sheetname=="履歴":
			rireki.contextMenuEntries(selection, entrynum)
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