#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。
from myrs import commons
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.sheet import CellFlags  # 定数

def getSectionName(controller, sheet, cell):  # 区画名を取得。
	"""
	M  |
	---
	C
	===========
	B  |D|E
	   | |
	-----------
	A
	"""
	rangeaddress = cell.getRangeAddress()  # セル範囲アドレスを取得。セルアドレスは不可。
	nonfreezedrow = 2  # 動く行の最上行のインデックス。
	contentcells = sheet[:, 2].queryContentCells(CellFlags.VALUE+CellFlags.STRING+CellFlags.FORMULA)  # 列インデックス2の数値か文字列か式の入っているセルに限定して抽出。空列は不可。
	emptyrow = contentcells.getRangeAddresses()[-1].EndRow + 1  # 最終行インデックス+1を取得。
	if len(sheet[0, :6].queryIntersection(rangeaddress)):  # メニューセルの時。
		return "M"
	elif len(sheet[nonfreezedrow:emptyrow, :8].queryIntersection(rangeaddress)):  # Dの左。
		return "B"	
	elif len(sheet[nonfreezedrow:emptyrow, 8:22].queryIntersection(rangeaddress)):  # チェック列の時。
		return "D"		
	elif len(sheet[nonfreezedrow:emptyrow, 22:].queryIntersection(rangeaddress)):  # Dの右。
		return "E"		
	elif len(sheet[emptyrow:, :].queryIntersection(rangeaddress)):  # まだデータのない行の時。
		return "A"	
	else:  # メニューセル以外の固定行の時。
		return "C"
def selectionChanged(controller, sheet, args):  # 矢印キーでセル移動した時も発火する。
	borders = args	
	selection = controller.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲がセルの時。矢印キーでセルを移動した時。マウスクリックハンドラから呼ばれると何回も発火するのでその対応。
		currenttableborder2 = selection.getPropertyValue("TableBorder2")  # 選択セルの枠線を取得。
		if all((currenttableborder2.TopLine.Color==currenttableborder2.LeftLine.Color==commons.COLORS["clearblue"],\
				currenttableborder2.RightLine.Color==currenttableborder2.BottomLine.Color==commons.COLORS["magenta"])):  # 枠線の色を確認。
			return  # すでに枠線が書いてあったら何もしない。
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		drowBorders(controller, sheet, selection, borders)	
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "", "血画を反映", ""),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, controller, sheet, target, args):  # マウスボタンを押した時。
	borders = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				drowBorders(controller, sheet, target, borders)
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
				section = getSectionName(controller, sheet, target)
				if section in ("M", "B", "D"):
					pass








					return False  # セル編集モードにしない。
	return True
def drowBorders(controller, sheet, cellrange, borders):  # ターゲットを交点とする行列全体の外枠線を描く。
	cell = cellrange[0, 0]  # セル範囲の左上端のセルで判断する。
	section = getSectionName(controller, sheet, cell)
	if section in ("A", "B", "D", "E"):
		noneline, tableborder2, topbottomtableborder, leftrighttableborder = borders	
		cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
		cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
		if cell.getPropertyValue("CellBackColor") in (-1, commons.COLORS["lightgreen"]):
			cellcursor = sheet.createCursorByRange(cellrange)  # targetをセル範囲とするセルカーサーを取得。
			cellcursor.expandToEntireColumns()  # 列全体を取得。
			if section=="D":
				cellcursor.setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
			cellcursor = sheet.createCursorByRange(cellrange)  # targetをセル範囲とするセルカーサーを再取得。
			cellcursor.expandToEntireRows()  # 行全体を取得。
			cellcursor.setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く。
			cellrange.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。	
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
