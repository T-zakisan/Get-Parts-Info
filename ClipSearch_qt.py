import sys
import time										#for sleep
import re											#正規表現
import pyodbc									#データベース
from pystray import Icon, MenuItem, Menu #タスクトレイ
from PIL import Image					#画像
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtGui
from pathlib import Path


# 環境ごとの設定 #########################################################
Path_DB     = ".\\部品番号台帳.accdb" 
Path_ICON   = r".\favicon.ico"
##########################################################################



# exe化したときのファイルパスを変更するヤツ ##########
def base_dir():
	if hasattr(sys, "_MEIPASS"):
		return Path(sys._MEIPASS)  # 実行ファイルで起動した場合
	else:
		return Path(".")  # python コマンドで起動した場合



# クラス ##########
class GetPartInfo(QMainWindow):
	
	# クラスの初期化 ##########
	def __init__( self, path_db, path_imIC ):
		super().__init__()	#クラスの初期化
		self.cur, self.conn = self.DBinit( path_db )	#pyodbc(データベース)の戻り値(接続先)
		self.im = Image.open( path_imIC )
		self.clipboard = QApplication.clipboard() 
		self.prev_clipboard_text = ""  # 前回のクリップボード内容
		self.clipboard_text = ""  #クリップボード内容
		self.createWindow()	#窓生成

		# 定期的なクリップボードチェック
		self.clipboard_timer = QTimer(self)
		self.clipboard_timer.timeout.connect(self.checkClipboard)
		self.clipboard_timer.start( 200 )  # 200msごとにチェック


	# クラス破壊 ##########
	def __del__( self ):
		self.conn.close()
		QApplication.quit()


	# 通知領域のアイコンから終了 ##########
	def Exit( self ):
		icon.stop()


	# データベースの初期設定 ##########
	def DBinit( self, path ):
		db_path = []
		db_path.append( path )	#データベースのパス
		con_str = (
        'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};' + 
				f'DBQ={ db_path[0] };'
		)	#データベースの接続設定
		conn =  pyodbc.connect( con_str )
		cur  =  conn.cursor()	#データベースへ接続
		return cur, conn


	# 窓の作成 ##########
	def createWindow( self ):
		self.setWindowTitle("Get Parts info")	#窓のタイトル
		self.setWindowFlags(
			Qt.Window |
			Qt.WindowStaysOnTopHint |
			Qt.CustomizeWindowHint
		)	#常前面表示
		self.setStyleSheet("self.MenuBar{color:#000000;}")			
		self.setGeometry( 0, 0, 320, 210)			#窓のサイズ
		self.setFixedSize( 320, 210 )

		# ドラッグ操作の設定
		self.mousePressEvent = self.startDrag
		self.mouseMoveEvent = self.performDrag

		# ラベルの表示 #####
		LABEL = [
			"部品番号：", "名　　称：", "型　式　：",
			"型　式２：", "型　式３：", "備　　考：" ]
		self.label = [ ["",""], ["",""], ["",""], ["",""], ["",""], ["",""] ]
		for ii in range( len( self.label ) ):
			for jj in range( len( self.label[ii] ) ):
				if jj == 0:
					self.label[ii][jj] = QLabel( LABEL[ii], self )	#ラベルを追加
					self.label[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 12 ) )	#書式
				else:
					self.label[ii][jj] = QLabel( "", self )	#ラベルを追加
					self.label[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 12, QtGui.QFont.Bold ) )
				self.label[ii][jj].adjustSize()
				self.label[ii][jj].setAlignment( Qt.AlignLeft )	#中央揃え
				self.label[ii][jj].move( 80*jj+10, 25*ii+5 )			#ラベルの位置
				if ii == 5 and jj == 1 :
					#備考
					self.label[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 11, QtGui.QFont.Bold ) )
					self.label[ii][jj].move( 80*0+10, 25*(ii+1)+5 )			#ラベルの位置
		# ボタン #####
		self.btn = QPushButton( '●', self )
		self.btn.setFont( QtGui.QFont( "BIZ UDゴシック", 11, QtGui.QFont.Bold ) )
		self.btn.setCheckable( True )						
		self.btn.setGeometry( 270, 0, 40, 25 )	#サイズと表示位置
		self.btn.pressed.connect( self.opsw )		#押したときの動作
		self.show()	#窓を表示


	# 窓の移動開始位置 ##########
	def startDrag(self, event):
			if event.button() == Qt.LeftButton:
					self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
					event.accept()

	# 窓の移動終了位置 ##########
	def performDrag(self, event):
			if event.buttons() == Qt.LeftButton:
					self.move(event.globalPos() - self.drag_position)
					event.accept()


	# 動作ボタンを押したときのイベント ##########
	def opsw( self ):
		if self.btn.text() == '●':
			#●(動作中)のときの処理
			self.btn.setText('■')	#■(停止)に変更
			self.btn.setStyleSheet("QPushButton{color:#c00000;}")	#文字色：真紅
			tmpClr = "#AAAAAA"	#ラベルの文字色：灰色(設定値のみ)
		else:
			#■(停止)のときの処理ZI010120
			self.btn.setText('●')	#●(動作中)に変更
			self.btn.setStyleSheet("QPushButton{color:#000000;}")	#文字色：黒
			tmpClr = "#000000"	#ラベルの文字色：黒(設定値のみ)
			self.clipboard.clear()  # クリップボードを初期化
		for ii in range( len( self.label ) ):
			for jj in range( len( self.label[ii] ) ):
				self.label[ii][jj].setStyleSheet("QLabel{color:"+tmpClr+";}")	#ラベルの文字色


	# Wクリックしたときのイベント ##########
	def mouseDoubleClickEvent( self, e ):
		self.opsw( )	


	# クリップボードの内容を監視し、変更があれば処理" ##########
	def checkClipboard(self):
		self.current_text = self.clipboard.text()
		if self.current_text != self.prev_clipboard_text and self.btn.text() == "●":
				self.prev_clipboard_text = self.current_text  # 変更を記録
				self.CheckCode( self.current_text )


	# キストの内容を正規表現でマッチした場合、データベースで検索し、通知へ ##########
	def CheckCode( self, code ):	
		#クリップボードのテキストを正規表現でマッチ
		ptn = '[a-zA-Z]{2}[0-9]{6}'	#部品番号のパタン ※組図以外
		rltsExRe	= re.search( ptn, code )	#クリップボードのテキストがptnと網羅的に一致

		#部品番号に一致がある場合の処理
		if rltsExRe:
			sql = "SELECT 品コード, 品名, 型式, 型式2, 型式3, 備考 FROM 部品番号 WHERE 品コード = ?"	#SQL文
			self.cur.execute( sql, rltsExRe[0] )	#クエリ
			rlts = self.cur.fetchone()	#結果を１つずつ取り出す ※部品番号が唯一のため成立
			outStr = ""
			tmp = ""
			for ii in range( len( self.label ) ):
				if rlts is None:
					if ii == 0:
						tmp = rltsExRe[0][0:7] + "No Data!"
						self.clipboard.setText( rltsExRe[0] ) 
					else:	tmp = ""
				else:
					if rlts[ ii ] is None : tmp = ""	#Noneで文字列結合するとエラー発動
					else				  			  : tmp = rlts[ ii ]
					if ii == 0: self.clipboard.setText( tmp )
					if ii == 5:
						outStr = ""
						for jj in range( 0, len( tmp ), 20 ):
							outStr += tmp[ jj: jj+20 ] + "\n"
					else: outStr = tmp
				self.label[ii][1].setText( outStr )
				self.label[ii][1].adjustSize()
				time.sleep( 0 )


	# 通知領域のアイコンからReadMe ##########
	def ReadMe( self, checked ):
		self.rd = ReadMe(  )
		self.rd.setWindowTitle("ReadMe")	#窓のタイトル
		self.rd.resize( 400, 560 )	#窓のサイズ
		self.rd.setFixedSize( 400, 560 )

		item = [
			[	"終了方法",
				"1.タスクバーのアイコン右クリ",
				"2.Exit"
			],[
				"窓移動",
				"1.窓のどこかをドラッグ(右クリっぱ)",
				"2.マウスで適度なところへ移動",
				"3.ドラッグを離す"
			],[
				"動作[●]/停止[■]モードの変更",
				"・窓をWクリック",
				"・窓右上のボタン[●/■]をクリック",
				"※[●]はクリップボードの自然な使用感を損なう"
			],[
				"仕様",
				"・常時前面表示",
				"・安易に消せない(通知領域で可能)",
				"・備考表示がチープ"
			],[
				"既知の問題点",
				"・高速にコピーを繰り返すと処理が落ちる.",
				"→焦んなって！"
			]
		]
		myStr = item.copy()
		cnt=0
		for ii in range( len( myStr ) ):
			col = 0
			for jj in range( len( myStr[ii] ) ):
				myStr[ii][jj] = QLabel( item[ii][jj], self.rd )
				if jj == 0:
					myStr[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 14, QtGui.QFont.Bold  ) )	#書式
					col = 5
				else:
					myStr[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 12  ) )	#書式
					col = 20
				myStr[ii][jj].setAlignment( Qt.AlignLeft )	#左揃え
				myStr[ii][jj].adjustSize()
				myStr[ii][jj].move( col, 25*cnt)	#ラベルの位置

				#セクションの最後でスペース
				add=1
				if jj==len( myStr[ii] )-1: add=2
				cnt+=add
		self.rd.show()


# クラス ##########
class ReadMe( QWidget ):
	def __init__( selt ):
		super().__init__( )


# main()メソッド ##########
if __name__ == '__main__':
	app = QApplication( sys.argv )	#Qtを初期化
	gpi = GetPartInfo(	path_db = Path_DB, 
											path_imIC = base_dir() / Path_ICON )
	menu = Menu( MenuItem( 'ReadMe', gpi.ReadMe ), MenuItem( 'Exit', gpi.Exit ) )
	icon = Icon( "icon", gpi.im, "部品番号チェッカ", menu = menu )
	icon.run()
	gpi.clipboard.setText('GetPartsInfo')
	sys.exit( gpi.Exit )
