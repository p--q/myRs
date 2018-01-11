#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。
def activeSpreadsheetChanged(sheet):  # シートがアクティブになった時。
	sheet["C1:F1"].setDataArray((("済をﾘｾｯﾄ", "", "血画を反映", ""),))  # よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
def singleClick(colors, controller, target):  # シングルクリックの時。
	
	# 行1のセルが結合しているかみる。
	

	celladdress = target.getCellAddress()
	rowindex = celladdress.Row  # ターゲットセルの行番号を取得。
	if rowindex >= controller.getSplitRow():  # 固定行ではない時。
		
		
		
# 		if ICHIRAN["leftendcolumn"]<celladdress.Column<ICHIRAN["rightendcolumn"]:
# 			pass
			# 縦罫線も書く。
			

	
	