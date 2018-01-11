#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os, sys
from types import ModuleType
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.util import XChangesListener
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.sheet import XActivationEventListener
from com.sun.star.sheet import CellFlags  # 定数
global XSCRIPTCONTEXT  # PyDevのエラー抑制用。
def enableRemoteDebugging(func):  # デバッグサーバーに接続したい関数やメソッドにつけるデコレーター。主にリスナーのメソッドのデバッグ目的。ただしマウスハンドラはフリーズするので直接pydevを書き込んだほうがよい。
	def wrapper(*args, **kwargs):
		frame = None
		doc = XSCRIPTCONTEXT.getDocument()
		if doc:  # ドキュメントが取得できた時
			frame = doc.getCurrentController().getFrame()  # ドキュメントのフレームを取得。
		else:
			currentframe = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()  # モードレスダイアログのときはドキュメントが取得できないので、モードレスダイアログのフレームからCreatorのフレームを取得する。
			frame = currentframe.getCreator()
		if frame:   
			import time
			indicator = frame.createStatusIndicator()  # フレームからステータスバーを取得する。
			maxrange = 2  # ステータスバーに表示するプログレスバーの目盛りの最大値。2秒ロスするが他に適当な告知手段が思いつかない。
			indicator.start("Trying to connect to the PyDev Debug Server for about 20 seconds.", maxrange)  # ステータスバーに表示する文字列とプログレスバーの目盛りを設定。
			t = 1  # プレグレスバーの初期値。
			while t<=maxrange:  # プログレスバーの最大値以下の間。
				indicator.setValue(t)  # プレグレスバーの位置を設定。
				time.sleep(1)  # 1秒待つ。
				t += 1  # プログレスバーの目盛りを増やす。
			indicator.end()  # reset()の前にend()しておかないと元に戻らない。
			indicator.reset()  # ここでリセットしておかないと例外が発生した時にリセットする機会がない。
		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  # デバッグサーバーを起動していた場合はここでブレークされる。import pydevdは時間がかかる。
		try:
			func(*args, **kwargs)  # Step Intoして中に入る。	
		except:
			import traceback; traceback.print_exc()  # これがないとPyDevのコンソールにトレースバックが表示されない。stderrToServer=Trueが必須。
	return wrapper
def macro(documentevent):  # 引数は文書のイベント駆動用。  
	colors = {"midori": 0x00FF00,\
			"pink": 0xFF00FF,\
			"kuro": 0x000000,\
			"ao": 0xFF0000,\
			"skyblue": 0xFFCC00,\
			"gray": 0xC0C0C0,\
			"aka": 0x0000FF}  # 色の16進数。	
	doc = documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	modulefolderpath = getModuleFolderPath(ctx, smgr, doc)  # 埋め込みモジュールフォルダへのURLを取得。
	modules = {}  # 埋め込みマクロのpythonpathフォルダのモジュールを入れる辞書。
	modulenames = "history", "ichiran", "karute", "keika", "schedule"  # 埋め込みpythonpathフォルダにあるモジュール名一覧。
	for modulename in modulenames:  # 埋め込みpythonpathフォルダにあるモジュールを辞書modulesに読み込む。
		modules[modulename] = load_module(simplefileaccess, "".join((modulefolderpath, "/", modulename, ".py")))
	sheets = doc.getSheets()
	sheet = sheets["一覧"]  # 一覧シートについて。

	doc.addChangesListener(ChangesListener(modules, doc))  # ChangesListener	
	controller = doc.getCurrentController()  # コントローラの取得。
	selectionchangelistener = SelectionChangeListener(modules, controller)  # SelectionChangeListener
	controller.addActivationEventListener(ActivationEventListener(modules, controller))  # ActivationEventListener
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(controller, colors, modules, selectionchangelistener))  # EnhancedMouseClickHandler
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(modules, ctx, smgr, doc))  # コントローラにContextMenuInterceptorを登録する。右クリックの時の対応。
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, modules, controller):
		self.controller = controller
		self.modules = modules
# 	@enableRemoteDebugging
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
		sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
		sheetname = sheet.getName()  # アクティブシート名を取得。
		if sheetname=="一覧":
			self.modules["ichiran"].activeSpreadsheetChanged(sheet)
	def disposing(self, eventobject):
		self.controller.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, controller, colors, modules, selectionchangelistener):
		self.controller = controller
		self.args = colors, modules, selectionchangelistener
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。
		colors, modules, selectionchangelistener = self.args
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			sheet = target.getSpreadsheet()
			sheetname = sheet.getName()  # アクティブシート名を取得。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時
				if sheetname=="一覧":
					modules["ichiran"].singleClick(colors, self.controller, target)
					
				# ここで罫線を引く
				
# 				controller.addSelectionChangeListener(self.selectionchangelistener)
			
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
# 				if controller.hasFrozenPanes():  # 表示→セルの固定、がされている時。
# 					splitrow = controller.getSplitRow()
# 					splitcolumn = controller.getSplitColumn()
				return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
	def mouseReleased(self, enhancedmouseevent):
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時		
				try:
					self.controller.removeSelectionChangeListener(self.selectionchangelistener)		
				except:
					pass
		return True
	def disposing(self, eventobject):
		self.controller.removeEnhancedMouseClickHandler(self)	
class SelectionChangeListener(unohelper.Base, XSelectionChangeListener):
	def __init__(self, modules, controller):
		self.controller = controller
	def selectionChanged(self, eventobject):
		controller = self.controller
		selection = controller.getSelection()  # 選択しているオブジェクトを取得。
	def disposing(self, eventobject):
		self.controller.removeSelectionChangeListener(self)		
class ChangesListener(unohelper.Base, XChangesListener):
	def __init__(self, modules, doc):
		self.subj = doc
	def changesOccurred(self, changesevent):
		changes = changesevent.Changes
		for change in changes:
			accessor = change.Accessor
			if accessor=="cell-change":  # セルの内容が変化した時。
				cell = change.ReplacedElement  # 変化したセルを取得。				
	def disposing(self, eventobject):
		self.doc.removeChangesListener(self)			
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, modules, ctx, smgr, doc):
		self.args = modules, getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
		consts, baseurl = self.args 
		controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
		global toBlue, toRed  # コンテクストメニューに割り当てる関数。
		toBlue, toRed = globalFunctionCreator(controller, consts)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
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
def globalFunctionCreator(controller, consts):
	colors = consts.COLORS
	target = controller.getSelection()  # 選択しているオブジェクトを取得。
	def toBlue():
		target.setPropertyValue("CellBackColor", colors["CellBackgroundColor"])  # 背景を青色にする。
	def toRed():
		target.setPropertyValue("CellBackColor", colors["CellRangeBackgroundColor"])  # 背景を赤色にする。
	return toBlue, toRed
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
def load_module(simplefileaccess, modulepath):  # modulepathのモジュールを取得。
	inputstream = simplefileaccess.openFileRead(modulepath)  # モジュールファイルからインプットストリームを取得。
	dummy, b = inputstream.readBytes([], inputstream.available())  # simplefileaccess.getSize(module_tdocurl)は0が返る。
	source = bytes(b).decode("utf-8")  # モジュールのソースをテキストで取得。
	mod = sys.modules.setdefault(modulepath, ModuleType(modulepath))  # 新規モジュールをsys.modulesに挿入。
	code = compile(source, modulepath, 'exec')  # urlを呼び出し元としてソースコードをコンパイルする。
	mod.__file__ = modulepath  # モジュールの__file__を設定。
	mod.__package__ = ''  # モジュールの__package__を設定。
	exec(code, mod.__dict__)  # 実行してモジュールの名前空間を取得。
	return mod
def getModuleFolderPath(ctx, smgr, doc):  # 埋め込みモジュールフォルダへのURLを取得。
	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
	tdocurl = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # ex. vnd.sun.star.tdoc:/1	
	return "/".join((tdocurl, "Scripts/python/pythonpath"))  # 開いているドキュメント内の埋め込みマクロフォルダへのパス。g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。	