<# ------------------------------------------------------------------------------------------
	ファイル名：子画面のあるダイアログボックス出力
	機能説明：	子画面を呼び出すボタンのあるダイアログボックスを出力
				各ダイアログボックスに任意の処理を実行するためのボタンを備える
------------------------------------------------------------------------------------------ #>

# 名前空間の読み込み
Using namespace System.Windows.Forms
Using namespace System.Drawing
# アセンブリ読み込み
Add-type -AssemblyName System.Windows.Forms
# エラー発生時は処理停止
$ErrorActionPreference = 'Stop'

Try
{
	# 【フォント】
	# オブジェクト作成
	$GlobalFont = New-Object Font('Meiryo UI', 10)

	# 【親画面】
	# オブジェクト作成
	$MainForm = New-Object Form
	# サイズを設定
	$MainForm.Size = New-Object Size(730, 400)
	# サイズ変更不可
	$MainForm.FormBorderStyle = 'FixedSingle'
	# 最大化ボタン非表示
	$MainForm.MaximizeBox = $False
	# 最小化ボタン非表示
	$MainForm.MinimizeBox = $False
	# 表示位置を設定
	$MainForm.StartPosition = 'CenterScreen'
	# タイトルを設定
	$MainForm.Text = 'MainForm'

	# 【親画面の入力用テキストボックス】
	# オブジェクト作成
	$MainInput = New-Object TextBox
	# オブジェクトの名前を設定
	$MainInput.Name = 'MainInputBox'
	# 表示位置を設定
	$MainInput.Location = New-Object Size(10, 60)
	# サイズを設定
	$MainInput.Size = New-Object Size(200, 20)
	# 最大入力値を設定（6桁まで）
	$MainInput.MaxLength = 6
	# 全角入力を制限
	$MainInput.ImeMode = [ImeMode]::Disable
	# デフォルト値を設定
	$MainInput.Text = '999999'
	# フォントを設定
	$MainInput.Font = $GlobalFont
	# 親画面に追加
	$MainForm.Controls.Add($MainInput)

	# 【親画面の表示用テキストボックス】
	# オブジェクト作成
	$OutputBox = New-Object TextBox
	# 表示位置を設定
	$OutputBox.Location = New-Object Size(10, 150)
	# サイズを設定
	$OutputBox.Size = New-Object Size(700, 180)
	# 複数行入力を可に設定
	$OutputBox.MultiLine = $True
	# 表示用のテキストボックスにスクロールバーを付与
	$OutputBox.ScrollBars = 'Vertical'
	# 手動入力を不可に設定
	$OutputBox.ReadOnly = $True
	# 表示するテキストを設定
	$OutputBox.Text = 'ログメッセージ出力エリア'
	# フォントを設定
	$OutputBox.Font = $GlobalFont
	# 親画面に追加
	$MainForm.Controls.Add($OutputBox)

	# 【親画面のラベル】
	# オブジェクト作成
	$MainLabel = New-Object Label
	# 表示位置を設定
	$MainLabel.Location = New-Object Point(10, 30)
	# サイズを設定
	$MainLabel.Size = New-Object Size(350, 20)
	# 表示するテキストを設定
	$MainLabel.Text = '主処理に必要な6桁の数値を入力してください'
	# フォントを設定
	$MainLabel.Font = $GlobalFont
	# 親画面に追加
	$MainForm.Controls.Add($MainLabel)

	# 【親画面のボタンA】
	# オブジェクト作成
	$MainButton_A = New-Object Button
	# 表示位置を設定
	$MainButton_A.Location = New-Object Size(450, 30)
	# サイズを設定
	$MainButton_A.Size = New-Object Size(250, 50)
	# フォントを設定
	$MainButton_A.Font = $GlobalFont
	# 表示するテキストを設定
	$MainButton_A.Text = '主処理を実行'
	# クリック時の処理を設定
	#$MainButton_A.Add_Click()
	# 親画面に追加
	$MainForm.Controls.Add($MainButton_A)

	# 【親画面のボタンB】
	# オブジェクト作成
	$MainButton_B = New-Object Button
	# 表示位置を設定
	$MainButton_B.Location = New-Object Size(450, 90)
	# サイズを設定
	$MainButton_B.Size = New-Object Size(250, 50)
	# フォントを設定
	$MainButton_B.Font = $GlobalFont
	# 表示するテキストを設定
	$MainButton_B.Text = '子画面を出力'
	# クリック時の処理を設定
	$MainButton_B.Add_Click({$SubForm.Show})
	# 親画面に追加
	$MainForm.Controls.Add($MainButton_B)

	# 【子画面】
	# オブジェクト作成
	$SubForm = New-Object Form
	# サイズを設定
	$SubForm.Size = New-Object Size(450, 200)
	# サイズ変更不可
	$SubForm.FormBorderStyle = 'FixedSingle'
	# 最大化ボタン非表示
	$SubForm.MaximizeBox = $False
	# 最小化ボタン非表示
	$SubForm.MinimizeBox = $False
	# 表示位置を設定
	$SubForm.StartPosition = 'Manual'
	# タイトルを設定
	$SubForm.Text = 'SubForm'
	# 親画面を設定
	$SubForm.Owner = $MainForm

	# 【子画面の入力用テキストボックス】
	# オブジェクト作成
	$SubInput = New-Object TextBox
	# オブジェクトの名前を設定
	$SubInput.Name = 'SubInputBox'
	# 表示位置を設定
	$SubInput.Location = New-Object Size(10, 60)
	# サイズを設定
	$SubInput.Size = New-Object Size(150, 20)
	# 最大入力値を設定（4桁まで）
	$SubInput.MaxLength = 4
	# 全角入力を制限
	$SubInput.ImeMode = [ImeMode]::Disable
	# フォントを設定
	$SubInput.Font = $GlobalFont
	# 子画面に追加
	$SubForm.Controls.Add($SubInput)

	# 【子画面のラベル】
	# オブジェクト作成
	$SubLabel = New-Object Label
	# 表示位置を設定
	$SubLabel.Location = New-Object Point(10, 30)
	# サイズを設定
	$SubLabel.Size = New-Object Size(350, 20)
	# 表示するテキストを設定
	$SubLabel.Text = '副処理に必要な4桁の数値を入力してください'
	# フォントを設定
	$SubLabel.Font = $GlobalFont
	# 子画面に追加
	$SubForm.Controls.Add($SubLabel)

	# 【子画面のボタン】
	# オブジェクト作成
	$SubButton = New-Object Button
	# 表示位置を設定
	$SubButton.Location = New-Object Size(230, 100)
	# サイズを設定
	$SubButton.Size = New-Object Size(200, 50)
	# フォントを設定
	$SubButton.Font = $GlobalFont
	# 表示するテキストを設定
	$SubButton.Text = '子画面の処理を実行'
	# クリック時の処理を設定
	#$SubButton.Add_Click()
	# 子画面に追加
	$SubForm.Controls.Add($SubButton)

	# 入力用テキストボックスの入力時イベントを設定
	$KeyPressEvent = {
		$PushKey = $_.KeyChar
		# 文字の入力を制限
		IF ((($PushKey -lt '0') -Or ($PushKey -gt '9')) -And -Not($PushKey -eq "`b")) {
			$_.Handled = $True
		}
	}
	$MainInput.Add_KeyPress($KeyPressEvent)
	$SubInput.Add_KeyPress($KeyPressEvent)

	# 親画面ボタンBのクリックイベントを設定
	$Click = {
		# 子画面の位置を設定
		$SubFormX = $MainForm.Location.X + 450
		$SubFormY = $MainForm.Location.Y - 25
		$SubForm.Location = New-Object Point($SubFormX, $SubFormY)
		# 子画面の入力用テキストボックスのデフォルト値を設定
		$SubInput.Text = '9999'
		# 子画面を表示
		$SubForm.ShowDialog()
	}
	$MainButton_B.Add_Click($Click)

	# 子画面のクロージングイベントを設定
	$Close = {
		$_.Cancel = $True
		$SubForm.Visible = $False
	}
	$SubForm.Add_Closing($Close)

	# 親画面の表示イベントを設定
	$Show = {
		$MainForm.Activate()
	}
	$MainForm.Add_Shown($Show)

	# 親画面を表示
	[Void] $MainForm.ShowDialog()
}
Finally # 最終処理
{
	# オブジェクト解放
	If ($Null -ne $GlobalFont) {$GlobalFont.Dispose()}
	If ($Null -ne $MainForm) {$MainForm.Dispose()}
}
