#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper  # オートメーションには必須(必須なのはuno)。
import os
from com.sun.star.awt import XEnhancedMouseClickHandler
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
try:
	from fordebugging import enableRemoteDebugging  # デバッグ用。マクロで実行した時。
except:
	pass
def macro(documentevent=None):  # 引数は文書のイベント駆動用。  
	doc = XSCRIPTCONTEXT.getDocument() if documentevent is None else documentevent.Source  # ドキュメントのモデルを取得。 
	ctx = XSCRIPTCONTEXT.getComponentContext()  # コンポーネントコンテクストの取得。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	
	
	
	
	controller = doc.getCurrentController()  # コントローラの取得。
	controller.addEnhancedMouseClickHandler(EnhancedMouseClickHandler())  # マウスハンドラをコントローラに設定。
	controller.registerContextMenuInterceptor(ContextMenuInterceptor(ctx, doc))  # コントローラにContextMenuInterceptorを登録する。
class EnhancedMouseClickHandler(unohelper.Base, XEnhancedMouseClickHandler): # マウスハンドラ
	def mousePressed(self, enhancedmouseevent):  # マウスボタンをクリックした時。ブーリアンを返さないといけない。
		target = enhancedmouseevent.Target  # ターゲットを取得。
		if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
			if enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 				import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)  
				if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
					sheet = target.getSpreadsheet()  # ターゲットがあるシートを取得。
					celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
					
					
					
					return False  # セル編集モードにしない。
		return True  # Falseを返すと右クリックメニューがでてこなくなる。
	def mouseReleased(self, enhancedmouseevent):  # ブーリアンを返さないといけない。
		return True  # Trueでイベントを次のハンドラに渡す。
	def disposing(self, eventobject):
		pass	
class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
	def __init__(self, ctx, doc):
		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
		name = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
		addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
		if name=="cell":  # セルのとき
			del contextmenu[:]  # contextmenu.clear()は不可。
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
		return EXECUTE_MODIFIED	

# class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
# 	def __init__(self, ctx, doc):
# 		self.baseurl = getBaseURL(ctx, doc)  # ScriptingURLのbaseurlを取得。
# 		global exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet   # ScriptingURLで呼び出す関数。オートメーションやAPSOでは不可。
# 		exportAsCSV, exportAsPDF, exportAsODS, SelectionToNewSheet = globalFunctionCreator(ctx, doc, sheet)  # クロージャーでScriptingURLで呼び出す関数に変数を渡す。
# 	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 引数はContextMenuExecuteEvent Struct。
# 		baseurl = self.baseurl  # ScriptingURLのbaseurlを取得。
# 		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # すでにあるコンテクストメニュー(アクショントリガーコンテナ)を取得。
# 		submenucontainer = contextmenu.createInstance("com.sun.star.ui.ActionTriggerContainer")  # サブメニューにするアクショントリガーコンテナをインスタンス化。
# 		addMenuentry(submenucontainer, "ActionTrigger", 0, {"Text": "Export as CSV...", "CommandURL": baseurl.format(exportAsCSV.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 1, {"Text": "Export as PDF...", "CommandURL": baseurl.format(exportAsPDF.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 2, {"Text": "Export as ODS...", "CommandURL": baseurl.format(exportAsODS.__name__)})  # サブメニューを挿入。引数のない関数名を渡す。
# 		addMenuentry(submenucontainer, "ActionTrigger", 3, {"Text": "Selection to New Sheet", "CommandURL": baseurl.format(SelectionToNewSheet.__name__)})
# 		addMenuentry(contextmenu, "ActionTrigger", 0, {"Text": "ExportAs", "SubContainer": submenucontainer})  # サブメニューを一番上に挿入。
# 		addMenuentry(contextmenu, "ActionTriggerSeparator", 1, {"SeparatorType": ActionTriggerSeparatorType.LINE})  # アクショントリガーコンテナのインデックス1にセパレーターを挿入。
		
		
		
# 		return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	i = 0  # インデックスを初期化する。
	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
		nonlocal i
		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
		i += 1  # インデックスを増やす。
	return addMenuentry
def addMenuentry(menucontainer, menutype, i, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
	menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
	[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
	menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
def getBaseURL(ctx, doc):	 # 埋め込みマクロ、オートメーション、マクロセレクターに対応してScriptingURLのbaseurlを返す。
	smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
	if modulepath.startswith(ucp):  # 埋め込みマクロの時。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
		filepath = modulepath.replace(ucp, "")  #  ucpを除去。
		transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
		transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
		contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
		macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
		location = "document"  # マクロの場所。	
	else:
		filepath = unohelper.fileUrlToSystemPath(modulepath) if modulepath.startswith("file://") else modulepath # オートメーションの時__file__はシステムパスだが、マクロセレクターから実行するとfileurlが返ってくる。
		pathsubstservice = smgr.createInstanceWithContext("com.sun.star.comp.framework.PathSubstitution", ctx)
		fileurl = pathsubstservice.substituteVariables("$(user)/Scripts/python", True)  # $(user)を変換する。fileurlが返ってくる。
		macrofolder =  unohelper.fileUrlToSystemPath(fileurl)  # fileurlをシステムパスに変換する。マイマクロフォルダへのパス。	
		location = "user"  # マクロの場所。
	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	
g_exportedScripts = macro, #マクロセレクターに限定表示させる関数をタプルで指定。		
if __name__ == "__main__":  # オートメーションで実行するとき
	from pythonpath.forautomation import automation
	try:
		from pythonpath.fordebugging import enableRemoteDebugging  # デバッグ用。
	except:
		pass
	XSCRIPTCONTEXT = automation()  # XSCRIPTCONTEXTを取得。	
	macro()  # マクロの実行。