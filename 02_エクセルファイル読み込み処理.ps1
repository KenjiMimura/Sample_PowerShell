<# ------------------------------------------------------------------------------------------
	ファイル名：エクセルファイル読み込み処理
	機能説明：	関数にエクセルファイルのパスとシート名を渡して、ファイルを読み込み
				エクセルファイルから取得したデータを指定したシート名と同名のテーブルに書き込み
------------------------------------------------------------------------------------------ #>

# 名前空間の読み込み
Using namespace System.Data
Using namespace System.Data.SqlClient
Using namespace System.Windows.Forms
Using namespace Microsoft.Office.Interop.Excel
Using namespace System.Runtime.InteropServices
# アセンブリ読み込み
Add-type -AssemblyName System.Windows.Forms
# エラー発生時は処理停止
$ErrorActionPreference = 'Stop'

# 定数定義
# DB関連
Set-Variable -Name DB_CONNECT -Value 'Data Source=SampleDataSource;Initial Catalog=TestDb;User ID=aaaaa;Password=zzzzz;' -Option Constant
Set-Variable -Name DB_TIMEOUT -Value 300 -Option Constant
# ログ関連
Set-Variable -Name LOG_PATH -Value '..\LOG\' -Option Constant
Set-Variable -Name LOG_NAME -Value 'Prefix_Log_' -Option Constant
Set-Variable -Name MSG_ST -Value '処理を開始します。' -Option Constant
Set-Variable -Name MSG_END -Value '処理を終了します。' -Option Constant
Set-Variable -Name MSG_ERR -Value 'エラーが発生しました。処理内容を確認して下さい。' -Option Constant
Set-Variable -Name MSG_TITLE -Value 'SampleTool' -Option Constant
# SQL関連
Set-Variable -Name SQL_GET_COL -Value "SELECT name FROM sys.columns WHERE object_id = object_id('?????') ORDER BY column_id;" -Option Constant
Set-Variable -Name REPLACE_WORD -Value '?????' -Option Constant

<# ------------------------------------------------------------------------------------------
	関数名：InsertExcelData（エクセルデータのDB入力処理）
	引数1：	InputFilePath（エクセルファイルパス）
	引数2：	InputSheet（シート名）
	戻り値：Bollean（正常=True/異常=False）
------------------------------------------------------------------------------------------ #>
Function InsertExcelData($InputFilePath, $InputSheet) {
	Try
	{
		# ログファイルの格納先を設定
		$FileName = $LOG_NAME + (Get-Date).ToString('yyyyMMdd') + '.log'
		$LogFullPath = $LOG_PATH + $FileName
		# ログファイル作成
		If ((Test-Path -Path $LogFullPath) -eq $False) {
			New-Item -Path $LOG_PATH -Name $FileName -ItemType 'file' > $Null
			Set-ItemProperty -Path $LogFullPath -Name IsReadOnly -Value $True
		}

		# 処理開始ログ出力
		If (-Not(OutputLog $MSG_ST $LogFullPath)) {Return $False}

		# DB接続
		$DBCon = New-Object SqlConnection($DB_CONNECT)
		$DBCon.Open()

		# コマンドオブジェクト作成
		$DBCmd = $DBCon.CreateCommand()
		# 実行SQL設定
		$DBCmd.CommandText = $SQL_GET_COL.Replace($REPLACE_WORD, $InputSheet)

		# データセットオブジェクト作成
		$Adapter = New-Object SqlDataAdapter($DBCmd)
		$DataSet = New-Object DataSet
		
		# SQL実行_出力先テーブルのカラム名を取得
		[Void]$Adapter.Fill($DataSet)

		# データテーブルオブジェクト作成
		$DataTable = New-Object Data.DataTable
		$DataSet.Tables[0].Rows.Name | Foreach-Object {[Void]$DataTable.Columns.Add($_)}

		# オブジェクト解放
		$DataSet.Dispose()
		$Adapter.Dispose()
		$DBCmd.Dispose()

		# Excelオブジェクト作成
		$Excel = New-Object -ComObject Excel.Application
		$Excel.Visible = $False
		$Excel.DisplayAlerts = $False

		# Excelファイル読み込み
		$xlBooks = $Excel.Workbooks;
		$TargetBook = $xlBooks.Open($InputFilePath)
		$xlSheets = $TargetBook.Worksheets;
		$TargetSheet = $xlSheets.Item($InputSheet)
		$xlCells = $TargetSheet.Cells;
		# 行と列の終端を取得
		$EndRow = $xlCells[1, 1].End([XlDirection]::xlDown).Row
		$EndCol = $xlCells[1, 1].End([XlDirection]::xlToRight).Column
		# ヘッダー行を除いたデータを取得
		$xlRange = $TargetSheet.Range($xlCells.Item(2, 1), $xlCells.Item($EndRow, $EndCol))
		$MasterData = $xlRange.Value(10)

		# Excelデータをデータテーブルにセット
		$ColCnt = 0
		$NewRow = $DataTable.NewRow()
		$MasterData | Foreach-Object {
			If ($Null -eq $_) {
				$NewRow[$ColCnt] = [DBNull]::Value
			} Else {
				$NewRow[$ColCnt] = $_
			}
			$ColCnt++
			# 終端カラムで新規のレコードを追加
			If ($ColCnt -eq $EndCol) {
				$DataTable.Rows.Add($NewRow)
				$ColCnt = 0
				$NewRow = $DataTable.NewRow()
			}
		}

		# バルクコピーオブジェクト作成
		$SqlBulk = New-Object SqlBulkCopy($DBCon)
		$SqlBulk.DestinationTableName = $InputSheet
		$SqlBulk.BulkCopyTimeout = $DB_TIMEOUT

		# データテーブルをDBに出力
		Foreach ($Col In $DataTable.Columns) {
			[Void]$SqlBulk.ColumnMappings.Add($Col.ColumnName, $Col.ColumnName)
		}
		$SqlBulk.WriteToServer($DataTable)

		# 処理終了ログ出力
		If (-Not(OutputLog $MSG_END $LogFullPath)) {Return $False}

		Return $True
	}
	Catch # 例外処理
	{
		# エラーログ出力
		[Void](OutputLog $MSG_ERR $LogFullPath)

		Return $False
	}
	Finally # 最終処理
	{
		# エクセル終了
		If ($Null -ne $xlRange) {[void][Marshal]::FinalReleaseComObject($xlRange)}
		If ($Null -ne $xlCells) {[void][Marshal]::FinalReleaseComObject($xlCells)}
		If ($Null -ne $TargetSheet) {[void][Marshal]::FinalReleaseComObject($TargetSheet)}
		If ($Null -ne $xlSheets) {[void][Marshal]::FinalReleaseComObject($xlSheets)}
		If ($Null -ne $TargetBook) {
			$TargetBook.Close()
			[void][Marshal]::FinalReleaseComObject($TargetBook)
		}
		If ($Null -ne $xlBooks) {[void][Marshal]::FinalReleaseComObject($xlBooks)}
		If ($Null -ne $Excel) {
			$Excel.Quit()
			[void][Marshal]::FinalReleaseComObject($Excel)
		}

		# オブジェクト解放
		If ($Null -ne $SqlBulk) {$SqlBulk.Dispose()}
		If ($Null -ne $DataTable) {$DataTable.Dispose()}
		If ($Null -ne $DataSet) {$DataSet.Dispose()}
		If ($Null -ne $Adapter) {$Adapter.Dispose()}
		If ($Null -ne $DBCmd) {$DBCmd.Dispose()}
		#　DB接続を切断
		If ($Null -ne $DBCon) {
			If ($DBCon.State -ne 'Closed') {$DBCon.Close()}
			$DBCon.Dispose()
		}
	}
}

<# ------------------------------------------------------------------------------------------
	関数名：OutputLog（ログファイル出力処理）
	引数1：	$BaseMSG（出力メッセージ）
	引数2：	$LogFullPath（ログファイルのフルパス）
	戻り値：Bollean（正常=True/異常=False）
------------------------------------------------------------------------------------------ #>
Function OutputLog($BaseMSG, $LogFullPath) {
	Try
	{
		# ログファイルに書き込み
		$LogMSG = (Get-Date).ToString() + ' ' + $BaseMSG
		Out-File -FilePath $LogFullPath -InputObject $LogMSG -Append -Force

		Return $True
	}
	Catch # 例外処理
	{
		# メッセージボックスを表示
		[Void][MessageBox]::Show($MSG_ERR, $MSG_TITLE, 'OK', 'Error')

		Return $False
	}
}
