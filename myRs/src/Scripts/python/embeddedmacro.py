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
from com.sun.star.sheet import XActivationEventListener


global XSCRIPTCONTEXT  # PyDevのエラー抑制用。
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	simplefileaccess = smgr.createInstanceWithContext("com.sun.star.ucb.SimpleFileAccess", ctx)  # SimpleFileAccess
	modulefolderpath = getModuleFolderPath(ctx, smgr, doc)  # 埋め込みモジュールフォルダへのURLを取得。
	consts = load_module(simplefileaccess, "/".join((modulefolderpath, "consts.py")))  # consts.pyをモジュールとして読み込む。
	controller.addActivationEventListener(ActivationEventListener(controller))  # ActivationEventListener
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler(controller))  # EnhancedMouseClickHandler

	
# 	セルを選択した時　罫線を引く




	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, smgr, doc, consts))  # コントローラにContextMenuInterceptorを登録する。右クリックの時の対応。
class ActivationEventListener(unohelper.Base, XActivationEventListener):
	def __init__(self, controller):  # subjはコントローラー。
		self.controller = controller
	def activeSpreadsheetChanged(self, activationevent):  # アクティブシートが変化した時に発火。
		sheet = activationevent.ActiveSheet  # アクティブになったシートを取得。
		sheetname = sheet.getName()  # アクティブシート名を取得。
		sheet["A1"].setString("ActiveSheetName: {}".format(sheetname))
	def disposing(self, eventobject):
		self.controller.removeActivationEventListener(self)	
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler):
	def __init__(self, controller):  # subjはコントローラー。
		self.controller = controller
	def mousePressed(self, enhancedmouseevent):  # セルをクリックした時に発火する。
		target = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			controller = self.controller
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
				if controller.hasFrozenPanes():  # 表示→セルの固定、がされている時。
					if 
					
					
					splitrow = controller.getSplitRow()
					splitcolumn = controller.getSplitColumn()
		
				
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					
					target.setString("R{}C{}".format(celladdress.Row, celladdress.Column))
					return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。		
		
		
		
		self._createLog(enhancedmouseevent, inspect.currentframe().f_code.co_name)
		return True
	def mouseReleased(self, enhancedmouseevent):
		self._createLog(enhancedmouseevent, inspect.currentframe().f_code.co_name)
		return True
	def disposing(self, eventobject):
		self.controller.removeEnhancedMouseClickHandler(self)
	def _createLog(self, enhancedmouseevent, methodname):
		dirpath, name = self.args
		target = enhancedmouseevent.Target
		target = getStringAddressFromCellRange(target) or target  # sourceがセル範囲の時は選択範囲の文字列アドレスを返す。
		clickcount = enhancedmouseevent.ClickCount
		filename = "_".join((name, methodname, "ClickCount", str(clickcount)))
		createLog(dirpath, filename, "Buttons: {}, ClickCount: {}, PopupTrigger {}, Modifiers: {}, Target: {}".format(enhancedmouseevent.Buttons, clickcount, enhancedmouseevent.PopupTrigger, enhancedmouseevent.Modifiers, target))	
		
		
		
		
		
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler): # マウスハンドラ
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					target.setString("R{}C{}".format(celladdress.Row, celladdress.Column))
					return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass	
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, smgr, doc, consts):
		self.args = consts, getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
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