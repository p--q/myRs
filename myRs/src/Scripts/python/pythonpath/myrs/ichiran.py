#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。
from myrs import consts
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "", "血画を反映", ""),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def singleClick(colors, controller, target):  # シングルクリックの時。
	
	# 行1のセルが結合しているかみる。
	

	celladdress = target.getCellAddress()
	rowindex = celladdress.Row  # ターゲットセルの行番号を取得。
	if rowindex >= controller.getSplitRow():  # 固定行ではない時。
		pass
		
		
		
# 		if ICHIRAN["leftendcolumn"]<celladdress.Column<ICHIRAN["rightendcolumn"]:
# 			pass
			# 縦罫線も書く。
			
def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, target, name):			
	if name=="cell":  # セルのとき
		del contextmenu[:]  # contextmenu.clear()は不可。
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # 引数のない関数名を渡す。
		elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # 引数のない関数名を渡す。
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
			
			

def contextMenuEntries(target, entrynum):
	colors = consts.COLORS
	if entrynum==1:
		target.setPropertyValue("CellBackColor", colors["ao"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["aka"]) 



# import unohelper
# import os
# from com.sun.star.ui import XContextMenuInterceptor
# from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
# from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
# from myrs import consts  # 相対インポートは不可。
# class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):  # コンテクストメニューのカスタマイズ。
# 	def __init__(self, ctx, smgr, doc):
# 		self.args = getBaseURL(ctx, smgr, doc)  # ScriptingURLのbaseurlを取得。
# 	def notifyContextMenuExecute(self, contextmenuexecuteevent):  # 右クリックで呼ばれる関数。contextmenuexecuteevent.ActionTriggerContainerを操作しないとコンテクストメニューが表示されない。
# # 		import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True) 
# 		baseurl = self.args 
# 		controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
# 		target = controller.getSelection()  # 選択しているオブジェクトを取得。
# 		contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
# 		name = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
# 		addMenuentry = menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 		if name=="cell":  # セルのとき
# 			del contextmenu[:]  # contextmenu.clear()は不可。
# 			if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
# 				addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format(toBlue.__name__)})  # 引数のない関数名を渡す。
# 			elif target.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
# 				addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format(toRed.__name__)})  # 引数のない関数名を渡す。
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 		elif name=="rowheader":  # 行ヘッダーのとき。
# 			del contextmenu[:]  # contextmenu.clear()は不可。
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
# 			addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
# 		elif name=="colheader":  # 列ヘッダーの時。
# 			pass  # contextmenuを操作しないとすべての項目が表示されない。
# 		elif name=="sheettab":  # シートタブの時。
# 			del contextmenu[:]  # contextmenu.clear()は不可。
# 			addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
# 		return EXECUTE_MODIFIED	  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。
# def toBlue():
# 	colors = consts.COLORS
# 	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
# 	target = doc.getCurrentSelection()
# 	target.setPropertyValue("CellBackColor", colors["ao"])  # 背景を青色にする。
# def toRed():
# 	colors = consts.COLORS
# 	doc = XSCRIPTCONTEXT.getDocument()  # ドキュメントのモデルを取得。 
# 	target = doc.getCurrentSelection()
# 	target.setPropertyValue("CellBackColor", colors["aka"])  # 背景を青色にする。
# def menuentryCreator(menucontainer):  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
# 	i = 0  # インデックスを初期化する。
# 	def addMenuentry(menutype, props):  # i: index, propsは辞書。menutypeはActionTriggerかActionTriggerSeparator。
# 		nonlocal i
# 		menuentry = menucontainer.createInstance("com.sun.star.ui.{}".format(menutype))  # ActionTriggerContainerからインスタンス化する。
# 		[menuentry.setPropertyValue(key, val) for key, val in props.items()]  #setPropertyValuesでは設定できない。エラーも出ない。
# 		menucontainer.insertByIndex(i, menuentry)  # submenucontainer[i]やsubmenucontainer[i:i]は不可。挿入以降のメニューコンテナの項目のインデックスは1増える。
# 		i += 1  # インデックスを増やす。
# 	return addMenuentry
# def getBaseURL(ctx, smgr, doc):	 # 埋め込みマクロのScriptingURLのbaseurlを返す。__file__はvnd.sun.star.tdoc:/4/Scripts/python/filename.pyというように返ってくる。
# 	modulepath = __file__  # ScriptingURLにするマクロがあるモジュールのパスを取得。ファイルのパスで場合分け。sys.path[0]は__main__の位置が返るので不可。
# 	ucp = "vnd.sun.star.tdoc:"  # 埋め込みマクロのucp。
# 	filepath = modulepath.replace(ucp, "")  #  ucpを除去。
# 	transientdocumentsdocumentcontentfactory = smgr.createInstanceWithContext("com.sun.star.frame.TransientDocumentsDocumentContentFactory", ctx)
# 	transientdocumentsdocumentcontent = transientdocumentsdocumentcontentfactory.createDocumentContent(doc)
# 	contentidentifierstring = transientdocumentsdocumentcontent.getIdentifier().getContentIdentifier()  # __file__の数値部分に該当。
# 	macrofolder = "{}/Scripts/python".format(contentidentifierstring.replace(ucp, ""))  #埋め込みマクロフォルダへのパス。	
# 	location = "document"  # マクロの場所。	
# 	relpath = os.path.relpath(filepath, start=macrofolder)  # マクロフォルダからの相対パスを取得。パス区切りがOS依存で返ってくる。
# 	return "vnd.sun.star.script:{}${}?language=Python&location={}".format(relpath.replace(os.sep, "|"), "{}", location)  # ScriptingURLのbaseurlを取得。Windowsのためにos.sepでパス区切りを置換。	