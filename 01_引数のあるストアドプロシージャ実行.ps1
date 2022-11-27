<# ------------------------------------------------------------------------------------------
	ファイル名：引数のあるストアドプロシージャ実行
	機能説明：	関数にストアドプロシージャの名称と引数を渡して、ストアドプロシージャを実行
------------------------------------------------------------------------------------------ #>

# 名前空間の読み込み
Using namespace System.Data
Using namespace System.Data.SqlClient
Using namespace System.Windows.Forms
# アセンブリ読み込み
Add-type -AssemblyName System.Windows.Forms
# エラー発生時は処理停止
$ErrorActionPreference = 'Stop'

# 定数定義
# DB関連
Set-Variable -Name DB_CONNECT -Value 'Data Source=SampleDataSource;Initial Catalog=TestDb;User ID=aaaaa;Password=zzzzz;' -Option Constant
Set-Variable -Name DB_TIMEOUT -Value 300 -Option Constant
# ストアドプロシージャ関連
Set-Variable -Name SP_Arg -Value 'TargetDateYMD' -Option Constant
Set-Variable -Name SP_RETURN -Value 'ReturnValue' -Option Constant
Set-Variable -Name SP_ERR_NUM -Value 'DBErrNum' -Option Constant
Set-Variable -Name SP_ERR_MSG -Value 'DBErrMsg' -Option Constant
Set-Variable -Name SP_ERR_LEN -Value 4000 -Option Constant
# ログ関連
Set-Variable -Name LOG_PATH -Value '..\LOG\' -Option Constant
Set-Variable -Name LOG_NAME -Value 'Prefix_Log_' -Option Constant
Set-Variable -Name MSG_ST -Value '処理を開始します。' -Option Constant
Set-Variable -Name MSG_END -Value '処理を終了します。' -Option Constant
Set-Variable -Name MSG_ERR -Value 'エラーが発生しました。処理内容を確認して下さい。' -Option Constant
Set-Variable -Name MSG_TITLE -Value 'SampleTool' -Option Constant

<# ------------------------------------------------------------------------------------------
	関数名：ExecProcedure（ストアドプロシージャ実行処理）
	引数1：	$ProcName（ストアドプロシージャの名称）
	引数2：	$ProcInput（ストアドプロシージャの引数）
	戻り値：Bollean（正常=True/異常=False）
------------------------------------------------------------------------------------------ #>
Function ExecProcedure($ProcName, $ProcInput) {
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
		$DBcmd.CommandType = [CommandType]::StoredProcedure
		$DBcmd.CommandText = $ProcName
		$DBCmd.CommandTimeout = $DB_TIMEOUT

		# ストアドプロシージャのパラメータを設定
		# 入力値
		[Void]$DBcmd.Parameters.Add($SP_Arg, [SqlDbType]::Int)
		$DBcmd.Parameters[$SP_Arg].Value = $ProcInput
		# 戻り値
		[Void]$DBcmd.Parameters.Add($SP_RETURN, [SqlDbType]::Int)
		$DBCmd.Parameters[$SP_RETURN].Direction = [ParameterDirection]::ReturnValue
		# エラーコード
		[Void]$DBcmd.Parameters.Add($SP_ERR_NUM, [SqlDbType]::Int)
		$DBCmd.Parameters[$SP_ERR_NUM].Direction = [ParameterDirection]::Output
		# エラーメッセージ
		[Void]$DBcmd.Parameters.Add($SP_ERR_MSG, [SqlDbType]::NVarChar, $SP_ERR_LEN)
		$DBCmd.Parameters[$SP_ERR_MSG].Direction = [ParameterDirection]::Output

		# ストアドプロシージャ実行
		$DataReader = $DBCmd.ExecuteReader()
		If ($DBCmd.Parameters[$SP_RETURN].Value -ne 0) {
			$DBErrCode = '[エラーコード：' + $DBCmd.Parameters[$SP_ERR_NUM].Value + '] '
			$ErrMSG = $MSG_ERR + "`r`n" + $DBErrCode + $DBCmd.Parameters[$SP_ERR_MSG].Value
			[Void](OutputLog $ErrMSG $LogFullPath)
			Return $False
		} Else {
			# 処理終了ログ出力
			If (-Not(OutputLog $MSG_END $LogFullPath)) {Return $False}
			
			Return $True
		}
	}
	Catch # 例外処理
	{
		# エラーログ出力
		[Void](OutputLog $MSG_ERR $LogFullPath)
		Return $False
	}
	Finally # 最終処理
	{
		# オブジェクト解放
		If ($DataReader.IsClosed -eq $False) {
			$DataReader.Close()
			$DataReader.Dispose()
		}	
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
