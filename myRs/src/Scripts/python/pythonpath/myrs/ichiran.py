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


def selectionChanged(controller, sheet, args):  # 矢印キーでセル移動sた時も発火する。しかしマウスで複数セルを選択すると枠線だらけになる。
	noneline, *borders = args
	selection = controller.getSelection()
	
	
	# 単独セルの時は左右、と上下に枠線を描く。複数セルの時は左、右、上、下とバラバラに線を描く、そうしないと複数セル範囲すべてに枠線が入ってしまう。
# 	if selection.supportsService("com.sun.star.sheet.SheetCell"):  # 選択範囲が単独セル範囲の時。
# 		
# 		
# 		pass
# 	
# 	elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲が複数セル範囲の時。
# 		pass
	
	
	
	
	
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
		cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
		cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
		cellcursor = sheet.createCursorByRange(selection)  # 選択範囲をセル範囲とするセルカーサーを取得。
		cellcursor.expandToEntireColumns()  # 列全体を取得。
		if cellcursor[0, 0].getIsMerged():  # セル範囲の先頭セルが結合セルの時。
			
			# 複数列ある時
			
			# 単独列の時。
			
			cellcursor.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。
			
		else:
			
			cellcursor[2:, :].setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。	
			
			
		cellcursor = sheet.createCursorByRange(selection)  # 選択範囲をセル範囲とするセルカーサーを再取得。
		cellcursor.expandToEntireRows()  # 行全体を取得。
		
		cellcursor.setPropertyValues(("TopBorder2", "BottomBorder2"), borders)  # 行の上下に枠線を引く。
		
# 		firstline, secondline = borders
# 		tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline, IsTopLineValid=True, IsBottomLineValid=True, IsLeftLineValid=True, IsRightLineValid=True)
# 		selection.setPropertyValue("TableBorder2", tableborder2)  # 消えた左右の枠線を引き直す。			
		
		selection.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 消えた左右の枠線を引き直す。


def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "", "血画を反映", ""),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def mousePressed(enhancedmouseevent, sheet, target, args):  # マウスボタンを押した時。
	noneline, *borders = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				
				# selectionChangedを無効にする。
				
				
				noneline, *borders = args
				if target.supportsService("com.sun.star.sheet.SheetCellRange"):  # ターゲットがセル範囲の時。
					cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
					cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
					cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを取得。
					cellcursor.expandToEntireColumns()  # 列全体を取得。
					if cellcursor[0, 0].getIsMerged():  # セル範囲の先頭セルが結合セルの時。
						cellcursor.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。
					else:
						cellcursor[2:, :].setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。	
					cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを再取得。
					cellcursor.expandToEntireRows()  # 行全体を取得。
					cellcursor.setPropertyValues(("TopBorder2", "BottomBorder2"), borders)  # 行の上下に枠線を引く。
					target.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 消えた左右の枠線を引き直す。				
				
				
				
# 				cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
# 				cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
# 				cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを取得。
# 				cellcursor.expandToEntireColumns()  # 列全体を取得。
# 				if cellcursor[0, 0].getIsMerged():  # セル範囲の先頭セルが結合セルの時。
# 					cellcursor.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # セルカーサーすべてのセルの左右に枠線を引く。
# 				else:
# 					cellcursor[2:, :].setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 行左右に枠線を引く。	
# 				cellcursor = sheet.createCursorByRange(target)  # targetをセル範囲とするセルカーサーを再取得。
# 				cellcursor.expandToEntireRows()  # 行全体を取得。
# 				cellcursor.setPropertyValues(("TopBorder2", "BottomBorder2"), borders)  # 上下に枠線を引く。
# 				target.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 消えた左右の枠線を引き直す。
				



				
	# 				controller.addSelectionChangeListener(self.selectionchangelistener)
			
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
# 			celladdress = target.getCellAddress()  # ターゲットのセルアドレスを取得。
# 				if controller.hasFrozenPanes():  # 表示→セルの固定、がされている時。
# 					splitrow = controller.getSplitRow()
# 					splitcolumn = controller.getSplitColumn()
				return False  # セル編集モードにしない。
	
	
	
	return True
	
	# 行1のセルが結合しているかみる。
	

# 	celladdress = target.getCellAddress()
# 	rowindex = celladdress.Row  # ターゲットセルの行番号を取得。
# 	if rowindex >= controller.getSplitRow():  # 固定行ではない時。
# 		pass
		
		
		

def mouseReleased(enhancedmouseevent, doc, sheet, target, args):  # マウスボタンを離した時。複数セルを選択した後でもtargetはセルしか入らない。
	
	# selectioncahnged()で枠線を描くときはここは不要。
	
	noneline, firstline, secondline = args
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ボタンのとき
		if target.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if enhancedmouseevent.ClickCount==1:  # シングルクリックの時。
				noneline, *borders = args
				selection = doc.getCurrentSelection()
				if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # セル範囲の時。
					cellcursor = sheet.createCursor()  # シートをセル範囲とするセルカーサーを取得。
# 					cellcursor.setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。
# 					cellcursor = sheet.createCursorByRange(selection)  # targetをセル範囲とするセルカーサーを取得。
# 					cellcursor.expandToEntireColumns()  # 列全体を取得。
# 					if cellcursor[0, 0].getIsMerged():  # セル範囲の先頭セルが結合セルの時。
# 						cellcursor.setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。
# 					else:
# 						cellcursor[2:, :].setPropertyValues(("LeftBorder2", "RightBorder2"), borders)  # 列の左右に枠線を引く。	
# 					cellcursor = sheet.createCursorByRange(selection)  # targetをセル範囲とするセルカーサーを再取得。
# 					cellcursor.expandToEntireRows()  # 行全体を取得。
# 					cellcursor.setPropertyValues(("TopBorder2", "BottomBorder2"), borders)  # 行の上下に枠線を引く。
					
					# tableborderを使わないと選択範囲内のセルすべてに外枠線が入る。
					
					tableborder2 = TableBorder2(TopLine=firstline, LeftLine=firstline, RightLine=secondline, BottomLine=secondline)
# 					selection.setPropertyValue("TableBorder2", tableborder2)  # 消えた左右の枠線を引き直す。				
					
					
					
			elif enhancedmouseevent.ClickCount==2:  # ダブルクリックの時
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
