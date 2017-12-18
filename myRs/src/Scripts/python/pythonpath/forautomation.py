#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import unohelper
import officehelper
import os, sys, glob
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.script.provider import XScriptContext  
from com.sun.star.document import MacroExecMode  # 定数
def automation():  # オートメーションのためにglobalに出すのはこの関数のみにする。
	try:
		ctx = officehelper.bootstrap()  # コンポーネントコンテクストの取得。
	except:
		print("Could not establish a connection with a running office.", file=sys.stderr)
		sys.exit()
	print("Connected to a running office ...")
	smgr = ctx.getServiceManager()  # サービスマネジャーの取得。
	print("Using {} {}".format(*getLOVersion(ctx, smgr)))  # LibreOfficeのバージョンを出力。
	os.chdir(os.path.join("..", "..", ".."))  # 呼び出したスクリプトから3つ上のプロジェクトディレクトリに移動。
	docs = glob.glob("*.ods")  # odsファイルを取得。
	if docs:
		ods = docs[0]  # 最初の一つのみ取得。
	else:
		print("There is no Calc document in {}".format(os.getcwd()))
		sys.exit()
	systempath = os.path.join(os.getcwd(), ods)  # odsファイルのフルパス。
	doc_fileurl = unohelper.systemPathToFileUrl(systempath)  # fileurlに変換。
	desktop = ctx.getByName('/singletons/com.sun.star.frame.theDesktop')  # デスクトップの取得。
	doc = getLoadedDocument(desktop, doc_fileurl)  # ドキュメントをすでに開いていたら取得。
	if doc is None:  # odsファイルが開いていない時。
		propertyvalues = PropertyValue(Name = "MacroExecutionMode", Value=MacroExecMode.ALWAYS_EXECUTE_NO_WARN),  # マクロを実行可能にする。
		desktop.loadComponentFromURL(doc_fileurl, "_blank", 0, propertyvalues) # ドキュメントを開く。ここでdocに代入してもドキュメントが開く前にmacro()が呼ばれてしまう。
		flg = True
		while flg:
			doc = getLoadedDocument(desktop, doc_fileurl)
			if doc is not None:  # odsファイルがデスクトップから取得出来た時。
				flg = False
	class ScriptContext(unohelper.Base, XScriptContext):
		def __init__(self, ctx):
			self.ctx = ctx
		def getComponentContext(self):
			return self.ctx
		def getDesktop(self):
			return desktop
		def getDocument(self):
			return doc
	return ScriptContext(ctx)  
def getLOVersion(ctx, smgr):  # LibreOfficeの名前とバージョンを返す。
	cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
	node = PropertyValue(Name = 'nodepath', Value = 'org.openoffice.Setup/Product' )  # share/registry/main.xcd内のノードパス。
	ca = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (node,))
	return ca.getPropertyValues(('ooName', 'ooSetupVersion'))  # LibreOfficeの名前とバージョンをタプルで返す。
def getLoadedDocument(desktop, doc_fileurl):  # ドキュメントをすでに開いていたらドキュメントモデルを返す。
	components = desktop.getComponents()  # ロードしているコンポーネントコレクションを取得。
	for component in components:  # 各コンポーネントについて。
		if hasattr(component, "getURL"):  # スタートモジュールではgetURL()はないためチェックする。
			if component.getURL()==doc_fileurl:  # fileurlが一致するとき、ドキュメントが開いているということ。
				return component  # componentがドキュメントモデル。		
			