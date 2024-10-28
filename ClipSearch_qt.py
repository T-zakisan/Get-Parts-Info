import sys
import pyperclip#							#クリップボード
import time										#for sleep
import re											#正規表現
import pyodbc									#データベース
import threading							#並列(別スレッド)処理
from pystray import Icon, MenuItem, Menu #タスクトレイ
from PIL import Image					#画像
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtGui
from pathlib import Path



# 環境ごとの設定 #########################################################
Path_DB     = ".\\testDB.accdb" 
Path_ICON   = r".\favicon.ico"
##########################################################################


def base_dir():
	if hasattr(sys, "_MEIPASS"):
		return Path(sys._MEIPASS)  # 実行ファイルで起動した場合
	else:
		return Path(".")  # python コマンドで起動した場合



class GetPartInfo(QMainWindow):

	def __init__( self, path_db, path_imIC ):
		super().__init__()	#クラスの初期化
		self.cur, self.conn = self.DBinit( path_db )	#pyodbc(データベース)の戻り値(接続先)
		self.im = Image.open( path_imIC )
		self.createWindow()	#窓生成


	def __del__( self ):
		self.conn.close()


	def Exit( self ):
		icon.stop()
		eventExit.set()


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


	def createWindow( self ):
		self.setWindowTitle("Get Parts info")	#窓のタイトル
		self.setWindowFlags( Qt.Window | Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint )	#常前面表示
		self.setGeometry( 0, 0, 320, 130)			#窓のサイズ
		self.setFixedSize( 320, 130 )

		LABEL = [
			"部品番号：",
			"名　　称：",
			"型　式　：",
			"型　式２：",
			"型　式３："
		]
		self.label = [ ["",""], ["",""], ["",""], ["",""], ["",""] ]
		for ii in range( len( self.label ) ):
			for jj in range( len( self.label[ii] ) ):
				if jj == 0:
					self.label[ii][jj] = QLabel( LABEL[ii], self )	#ラベルを追加
					self.label[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 12 ) )	#書式
					# self.label[ii][jj].setText( LABEL[ii] )
				else:
					self.label[ii][jj] = QLabel( "", self )	#ラベルを追加
					self.label[ii][jj].setFont( QtGui.QFont( "BIZ UDゴシック", 12, QtGui.QFont.Bold ) )
				self.label[ii][jj].adjustSize()
				self.label[ii][jj].setAlignment( Qt.AlignLeft )	#中央揃え
				self.label[ii][jj].move( 80*jj+10, 25*ii+5 )			#ラベルの位置

		self.show()	#窓を表示


	def CheckCB( self ):
		pyperclip.copy( '' )	#クリップボードを初期化
		while not( eventExit.is_set() ):
			time.sleep( 0.2 )	#適当なDelay
			code = pyperclip.waitForPaste()	#クリップボードの更新
			#code = pyperclip.waitForNewPaste()	#クリップボードの更新
			if( code != "" and isinstance( code, str ) == True ):	#クリップボードの内容が空でなく、文字列の場合
				GetPartInfo.CheckCode( self, code )	#クリップボードの内容をチェック


	#テキストの内容を正規表現でマッチした場合、データベースで検索し、通知へ
	def CheckCode( self, code ):	
		#クリップボードのテキストを正規表現でマッチ
		ptn = '[A-Z]{2}[0-9]{6}'	#部品番号のパタン ※組図以外
		rltsExRe	= re.findall( ptn, code )	#クリップボードのテキストがptnと網羅的に一致

		#部品番号に一致がある場合の処理
		if rltsExRe:
			for rltExRe in rltsExRe:	#複数ある部品番号でくり返し
				# pyperclip.copy( '' )
				sql = "SELECT PartNo, 品名, 型式, 型式2, 型式3 FROM ttt WHERE PartNo = ?"	#SQL文
				self.cur.execute( sql, [rltExRe] )	#クエリ
				rlts = self.cur.fetchone()	#結果を１つずつ取り出す ※部品番号が唯一のため成立
				tmp = ""
				if rlts is None:	#データベースにない場合の処理
					#通知(NoData)
					for ii in range( 5 ):
						if ii == 0:
							tmp = "No Data!"
						else:
							tmp = ""
						self.label[ii][1].setText( tmp )
						self.label[ii][1].adjustSize()
					time.sleep( 0 )
				else:	#データベースにヒットした場合の処理
					for ii, rr in enumerate( rlts ):
							if ii == 0 : pyperclip.copy( rr )
							if rr is None : tmp = ""	#Noneで文字列結合するとエラー発動
							else					: tmp = rr
							self.label[ii][1].setText( tmp )
							self.label[ii][1].adjustSize()
					time.sleep( 0 )
				break	#forから抜け出し


	def ReadMe( self, checked ):
		self.rd = ReadMe(  )
		self.rd.setWindowTitle("ReadMe")	#窓のタイトル
		self.rd.resize( 480, 320 )	#窓のサイズ
		self.rd.setFixedSize( 480, 320 )

		item = [
			[	"終了方法",
				"1.タスクバーのアイコン右クリ",
				"2.Exit"
			],[
				"窓移動",
				"1.窓の上端をドラッグ",
				"2.マウスで適度なところへ移動",
				"3.ドラッグを離す"
			],[
				"既知の問題点",
				"・高速にコピーをくり返した場合に落ちる.",
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
				myStr[ii][jj].move( col, 25*cnt)#25*jj+100*ii )			#ラベルの位置

				#セクションの最後でスペース
				add=1
				if jj==len( myStr[ii] )-1: add=2
				cnt+=add
		self.rd.show()



class ReadMe( QWidget ):
	def __init__( selt ):
		super().__init__( )



#main()メソッド
if __name__ == '__main__':
	app = QApplication( sys.argv )	#Qtを初期化
	gpi = GetPartInfo(	path_db = Path_DB, 
											path_imIC = base_dir() / Path_ICON )
	eventExit = threading.Event()
	menu = Menu( MenuItem( 'ReadMe', gpi.ReadMe ), MenuItem( 'Exit', gpi.Exit ) )
	icon = Icon( "icon", gpi.im, "部品番号チェッカ", menu = menu )

	CheckTrd = threading.Thread( target = gpi.CheckCB )
	CheckTrd.start()
	icon.run()
	pyperclip.copy( 'Exit GetPartsInfo' )	#Exitのための処理
	CheckTrd.join()
	sys.exit( gpi.Exit )
