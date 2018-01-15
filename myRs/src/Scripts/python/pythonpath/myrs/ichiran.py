#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。
from myrs import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.table import BorderLine2  # Struct
from com.sun.star.table import TableBorder2  # Struct

# def defineArea(sheet):
# 	pass


def selectionChanged(controller, sheet, args):  # 矢印キーでセル移動した時も発火する。
	noneline, firstline, secondline = args
	selection = controller.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):
		return  # 選択範囲がセルの時は何もしない。
	elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル以外のセル範囲の時。
		cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
		cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
		cellcursor = sheet.createCursorByRange(selection)  # 選択範囲をセル範囲とするセルカーサーを取得。
		cellcursor.expandToEntireColumns()  # 列全体を取得。
		if cellcursor[0, 0].getIsMerged():  # セルカーサーの先頭セルが結合セルの時。
			cellcursor[:, 0].setPropertyValue("LeftBorder2", firstline)  # セルカーサーの左端列の左に枠線を引く。
			cellcursor[:, -1:].setPropertyValue("RightBorder2", secondline)  # セルカーサーの右端列の右に枠線を引く。
		cellcursor = sheet.createCursorByRange(selection)  # 選択範囲をセル範囲とするセルカーサーを再取得。
		cellcursor.expandToEntireRows()  # 行全体を取得。
		cellcursor[0, :].setPropertyValue("TopBorder2", firstline)  # セルカーサーの最上行の上に枠線を引く。
		cellcursor[-1:, :].setPropertyValue("BottomBorder2", secondline)  # セルカーサーの最下行の下に枠線を引く。
		tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
		selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "", "血画を反映", ""),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, sheet, target, args):  # マウスボタンを押した時。
	noneline, firstline, secondline = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
					
					cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
					cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
					cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを取得。
					cellcursor.expandToEntireColumns()  # 列全体を取得。
					if cellcursor[0, 0].getIsMerged():  # セル範囲の先頭セルが結合セルの時。
						cellcursor.setPropertyValues(("LeftBorder2", "RightBorder2"), (firstline, secondline))  # 列の左右に枠線を引く。
					else:
						cellcursor[2:, :].setPropertyValues(("LeftBorder2", "RightBorder2"), (firstline, secondline))  # 列の左右に枠線を引く。	
					cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを再取得。
					cellcursor.expandToEntireRows()  # 行全体を取得。
					cellcursor.setPropertyValues(("TopBorder2", "BottomBorder2"), (firstline, secondline))  # 行の上下に枠線を引く。
					tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
					target.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	


			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 			celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
# 				if controller.hasFrozenPanes():  # 表示→セルの固定、がされている時。
# 					splitrow = controller.getSplitRow()
# 					splitcolumn = controller.getSplitColumn()
				return False  # セル編集モードにしない。
	
	
	
	return True
def notifycontextmenuexecute(addMenuentry, baseurl, contextmenu, controller, contextmenuname):  # 右クリックメニュー。			
	if contextmenuname=="cell":  # セルのとき
		selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
		del contextmenu[:]  # contextmenu.clear()は不可。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
			addMenuentry("ActionTrigger", {"Text": "To blue", "CommandURL": baseurl.format("entry1")})  # listeners.pyの関数名を指定する。
		elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 連続した複数セルの時。
			addMenuentry("ActionTrigger", {"Text": "To red", "CommandURL": baseurl.format("entry2")})  # listeners.pyの関数名を指定する。
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
	elif contextmenuname=="rowheader":  # 行ヘッダーのとき。
		del contextmenu[:]  # contextmenu.clear()は不可。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Cut"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Copy"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Paste"})
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:InsertRowsBefore"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:DeleteRows"}) 
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass  # contextmenuを操作しないとすべての項目が表示されない。
	elif contextmenuname=="sheettab":  # シートタブの時。
		del contextmenu[:]  # contextmenu.clear()は不可。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
def contextMenuEntries(target, entrynum):  # コンテクストメニュー番号の処理を振り分ける。
	colors = commons.COLORS
	if entrynum==1:
		target.setPropertyValue("CellBackColor", colors["blue"])  # 背景を青色にする。
	elif entrynum==2:
		target.setPropertyValue("CellBackColor", colors["red"]) 
