#################################################################################
# 処理名　｜ResizeImageTool
# 機能　　｜画像ファイルのリサイズツール
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜なし
#################################################################################
# 設定
# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み（フォーム用）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = 'Stop'
# 定数
[System.String]$c_config_file = 'setup.ini'
# エラーコード enum設定
Add-Type -TypeDefinition @"
    public enum MESSAGECODE {
        Successful = 0,
        Abend,
        Cancel,
        Info_LoadedSettingfile,
        Info_WebpSkipBatchOcrKeywordcount,
        Info_ComplateConvertWebp,
        Info_WebpSkipConvertWebp,
        Info_WebpSkipResizeImage,
        Info_WebpSkipSettingSizeIsBigger,
        Info_ComplateResizeImage,
        Confirm_ExecutionTool,
        Confirm_OcrResult,
        Confirm_ResizeImages,
        Error_NotCore,
        Error_NotSupportedVersion,
        Error_NotWindows,
        Error_LoadingSettingfile,
        Error_NotExistsTargetpath,
        Error_EmptyTargetfolder,
        Error_EmptyResizeValue,
        Error_ZeroResizeValue,
        Error_NotIntResizeValue,
        Error_MaxRetries,
        Error_EmptyOcrExepath,
        Error_EmptyOcrTemppath,
        Error_EmptyOcrSearchKeyword,
        Error_NotExistsOcrExepath,
        Error_NotExistsOcrTemppath,
        Error_CopyTempfile,
        Error_ExecuteTesseractOcr,
        Error_RemoveTempfile,
        Error_OverSizeForMonitor,
        Error_ExecuteResize,
        Error_CreateResizeFolder,
        Error_ResizefileSave,
        Error_ChangeExtension,
        Error_ConvertWebp,
        Error_RemoveFile
    }
"@

### DEBUG ###
Set-Variable -Name "DEBUG_ON" -Value $false -Option Constant

### Function --- 開始 --->
#################################################################################
# 処理名　｜RemoveDoubleQuotes
# 機能　　｜先頭桁と最終桁にあるダブルクォーテーションを削除
#--------------------------------------------------------------------------------
# 戻り値　｜String（削除後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function RemoveDoubleQuotes {
    param (
        [System.String]$target_str
    )
    [System.String]$removed_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq '"') -and
            ($target_str.Substring($target_str.Length - 1, 1) -eq '"')) {
            # 先頭桁と最終桁のダブルクォーテーション削除
            $removed_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RemoveDoubleQuotes: target_str  [${target_str}]"
        Write-Host "                             removed_str [${removed_str}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $removed_str
}

#################################################################################
# 処理名　｜VerificationExecutionEnv
# 機能　　｜実行環境の検証
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜なし
#################################################################################
Function VerificationExecutionEnv {
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_message = ''

    # 環境情報を取得
    [System.Collections.Hashtable]$powershell_ver = $PSVersionTable

    # 環境の判定：Coreではない場合
    if ($powershell_ver.PSEdition -ne 'Core') {
        $messagecode = [MESSAGECODE]::Error_NotCore
        $messagecode_message = RetrieveMessage $messagecode
        Write-Host $messagecode_message -ForegroundColor DarkRed
    }
    # 環境の判定：メジャーバージョンが7より小さい場合
    elseif ($powershell_ver.PSVersion.Major -lt 7) {
        $messagecode = [MESSAGECODE]::Error_NotSupportedVersion
        $messagecode_message = RetrieveMessage $messagecode
        Write-Host $messagecode_message -ForegroundColor DarkRed
    }
    # 環境の判定：Windows OSではない場合
    elseif (-Not($IsWindows)) {
        $messagecode = [MESSAGECODE]::Error_NotWindows
        $messagecode_message = RetrieveMessage $messagecode
        Write-Host $messagecode_message -ForegroundColor DarkRed
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function VerificationExecutionEnv: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
#################################################################################
# 処理名　｜AcquisitionFormsize
# 機能　　｜Windowsフォーム用のサイズをモニターサイズから除算で設定
#--------------------------------------------------------------------------------
# 戻り値　｜String[]（変換後のサイズ：1要素目 横サイズ、2要素目 縦サイズ）
# 引数　　｜divisor: 除数（モニターサイズから除算するため）
#################################################################################
Function AcquisitionFormsize {
    param (
        [System.UInt32]$divisor
    )
    # 現在のモニターサイズを取得
    [Microsoft.Management.Infrastructure.CimInstance]$graphics_info = (Get-CimInstance -ClassName Win32_VideoController)
    [System.UInt32]$width = $graphics_info.CurrentHorizontalResolution
    [System.UInt32]$height = $graphics_info.CurrentVerticalResolution

    # モニターのサイズから除数で割る
    [System.UInt32]$form_width = $width / $divisor
    [System.UInt32]$form_height = $height / $divisor
    
    [System.UInt32[]]$form_size = @($form_width, $form_height)

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function AcquisitionFormsize: form_size [${form_size}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $form_size
}

#################################################################################
# 処理名　｜ConfirmYesno
# 機能　　｜YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno {
    param (
        [System.String]$prompt_message,
        [System.String]$prompt_title='実行前の確認'
    )

    # 除数「6」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(6)

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(($form_size[0] * 0.016), ($form_size[1] * 0.030))
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(($form_size[0] * 0.0156),($form_size[1] * 0.0285))
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.07))
    $label.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label.Text = $prompt_message
    $label.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = 'OK'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # フォーム表示
    [System.Boolean]$is_selected = ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: is_selected [${is_selected}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $is_selected
}

#################################################################################
# 処理名　｜ValidateInputValues
# 機能　　｜入力値の検証
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜setting_parameters[]
# 　　　　｜ - 項目01 作業フォルダー
# 　　　　｜ - 項目02 リサイズの横サイズ
# 　　　　｜ - 項目03 リサイズの縦サイズ
#################################################################################
Function ValidateInputValues {
    param (
        [System.String[]]$setting_parameters
    )
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful

    # メッセージボックス用
    [System.String]$messagebox_title = ''
    [System.String]$messagebox_messages = ''
    [System.String]$append_message = ''

    # 作業フォルダー
    #   入力チェック
    if ($setting_parameters[0] -eq '') {
        $messagecode = [MESSAGECODE]::Error_EmptyTargetfolder
        $messagebox_messages = RetrieveMessage $messagecode
        $messagebox_title = '入力チェック'
        ShowMessagebox $messagebox_messages $messagebox_title
    }
    #   存在チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        if (-Not(Test-Path $setting_parameters[0])) {
            $messagecode = [MESSAGECODE]::Error_NotExistsTargetpath
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
            "作業フォルダー: [$($setting_parameters[0])]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $messagebox_messages = RetrieveMessage $messagecode $append_message
            $messagebox_title = '存在チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    # リサイズの縦横サイズ
    #   入力チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        # 未入力の場合
        if (($setting_parameters[1] -eq '') -and
            ($setting_parameters[2] -eq '')) {
            $messagecode = [MESSAGECODE]::Error_EmptyResizeValue
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
        # Zero（0）が入力された場合
        if (($setting_parameters[1] -eq 0) -or
            ($setting_parameters[2] -eq 0)) {
            $messagecode = [MESSAGECODE]::Error_ZeroResizeValue
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }
    #   数値チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        # Witdh
        If ($setting_parameters[1] -ne '') {
            if (-Not([UINt32]::TryParse($setting_parameters[1],[ref]$null))) {
                $messagecode = [MESSAGECODE]::Error_NotIntResizeValue
                $sbtemp=New-Object System.Text.StringBuilder
                @("`r`n",`
                "リサイズする横サイズ（width） : [$($setting_parameters[1])]`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $append_message = $sbtemp.ToString()
                $messagebox_messages = RetrieveMessage $messagecode $append_message
                $messagebox_title = '数値チェック'
                ShowMessagebox $messagebox_messages $messagebox_title
            }
        }
        # Height
        If ($setting_parameters[2] -ne '') {
            if (-Not([UINt32]::TryParse($setting_parameters[2],[ref]$null))) {
                $messagecode = [MESSAGECODE]::Error_NotIntResizeValue
                $sbtemp=New-Object System.Text.StringBuilder
                @("`r`n",`
                "リサイズする縦サイズ（height）: [$($setting_parameters[2])]`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $append_message = $sbtemp.ToString()
                $messagebox_messages = RetrieveMessage $messagecode $append_message
                $messagebox_title = '数値チェック'
                ShowMessagebox $messagebox_messages $messagebox_title
            }
        }
    }

    #   矛盾チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        [System.UInt32[]]$monitor_size = AcquisitionFormsize(1)
        # モニターサイズより大きな値が設定されている場合はエラーとする
        If (($setting_parameters[1] -eq '') -and
            ($setting_parameters[1] -gt $monitor_size[0])) {
            $messagecode = [MESSAGECODE]::Error_OverSizeForMonitor
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
              "リサイズする横サイズ（width）: [$($setting_parameters[1])]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $messagebox_messages = RetrieveMessage $messagecode $append_message
            $messagebox_title = '矛盾チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
        If (($setting_parameters[2] -eq '') -and
            ($setting_parameters[2] -gt $monitor_size[1])) {
            $messagecode = [MESSAGECODE]::Error_OverSizeForMonitor
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
              "リサイズする縦サイズ（height）: [$($setting_parameters[2])]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $messagebox_messages = RetrieveMessage $messagecode $append_message
            $messagebox_title = '矛盾チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConfirmYesno: return [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　｜SettingInputValues
# 機能　　｜入力情報の設定（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Object[]
# 　　　　｜ - 項目01 作業フォルダー    : 画面での設定値 - 画像ファイルをチェックし変換する作業対象のフォルダー
# 　　　　｜ - 項目02 リサイズの横サイズ: 画面での設定値 - リサイズ後の横サイズ（px）
# 　　　　｜ - 項目03 リサイズの縦サイズ: 画面での設定値 - リサイズ後の縦サイズ（px）
# 引数　　｜function_parameters[]
# 　　　　｜ - 項目01 ツール実行場所    : ツールの実行場所
# 　　　　｜ - 項目02 作業フォルダー    : 初期表示用の値 - 画像ファイルをチェックし変換する作業対象のフォルダー
# 　　　　｜ - 項目03 リサイズの横サイズ: 初期表示用の値 - リサイズ後の横サイズ（px）
# 　　　　｜ - 項目04 リサイズの縦サイズ: 初期表示用の値 - リサイズ後の縦サイズ（px）
#################################################################################
Function SettingInputValues {
    param (
        [System.Object[]]$function_parameters
    )

    # 除数「3」で割った値をフォームサイズとする
    [System.UInt32[]]$form_size = AcquisitionFormsize(3)

    # フォームの作成
    [System.String]$prompt_title = '実行前の設定'
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = $prompt_title
    $form.Size = New-Object System.Drawing.Size($form_size[0],$form_size[1])
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("$($function_parameters[0])\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.FormBorderStyle = 'FixedSingle'

    # 作業フォルダー - ラベル作成
    [System.Windows.Forms.Label]$label_work_folder = New-Object System.Windows.Forms.Label
    $label_work_folder.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.070))
    $label_work_folder.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_work_folder.Text = '参照するフォルダーを指定してください。'
    $label_work_folder.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 作業フォルダー - テキストボックスの作成
    [System.Windows.Forms.TextBox]$textbox_work_folder = New-Object System.Windows.Forms.TextBox
    $textbox_work_folder.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.175))
    $textbox_work_folder.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75), 0)
    $textbox_work_folder.Text = $function_parameters[1]
    $textbox_work_folder.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 作業フォルダー - 参照ボタンの作成
    [System.Windows.Forms.Button]$btnRefer = New-Object System.Windows.Forms.Button
    $btnRefer.Location = New-Object System.Drawing.Point(($form_size[0] * 0.820), ($form_size[1] * 0.175))
    $btnRefer.Size = New-Object System.Drawing.Size(75,25)
    $btnRefer.Text = '参照'

    # 横のサイズ - ラベル作成
    [System.Windows.Forms.Label]$label_resize_width = New-Object System.Windows.Forms.Label
    $label_resize_width.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.28))
    $label_resize_width.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_resize_width.Text = 'リサイズする横のサイズを指定（px）してください。'
    $label_resize_width.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 横のサイズ - テキストボックスの作成
    [System.Windows.Forms.TextBox]$textbox_resize_width = New-Object System.Windows.Forms.TextBox
    $textbox_resize_width.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.385))
    $textbox_resize_width.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75), 0)
    $textbox_resize_width.Text = $function_parameters[2]
    $textbox_resize_width.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 縦のサイズ - ラベル作成
    [System.Windows.Forms.Label]$label_resize_height = New-Object System.Windows.Forms.Label
    $label_resize_height.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04),($form_size[1] * 0.49))
    $label_resize_height.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75),($form_size[1] * 0.075))
    $label_resize_height.Text = 'リサイズする縦のサイズを指定（px）してください。'
    $label_resize_height.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # 縦のサイズ - テキストボックスの作成
    [System.Windows.Forms.TextBox]$textbox_resize_height = New-Object System.Windows.Forms.TextBox
    $textbox_resize_height.Location = New-Object System.Drawing.Point(($form_size[0] * 0.04), ($form_size[1] * 0.595))
    $textbox_resize_height.Size = New-Object System.Drawing.Size(($form_size[0] * 0.75), 0)
    $textbox_resize_height.Text = $function_parameters[3]
    $textbox_resize_height.Font = New-Object System.Drawing.Font('ＭＳ ゴシック',11)

    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(($form_size[0] - 205), ($form_size[1] - 90))
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = '次へ'
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(($form_size[0] - 115), ($form_size[1] - 90))
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = 'キャンセル'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel

    # フォームに紐づけ
    $form.Controls.Add($label_work_folder)
    $form.Controls.Add($textbox_work_folder)
    $form.Controls.Add($label_resize_width)
    $form.Controls.Add($textbox_resize_width)
    $form.Controls.Add($label_resize_height)
    $form.Controls.Add($textbox_resize_height)
    $form.Controls.Add($btnRefer)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)

    # 参照ボタンの処理
    $btnRefer.add_click{
        #ダイアログを表示しファイルを選択する
        $folder_dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        if($folder_dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            $textbox_work_folder.Text = $folder_dialog.SelectedPath
        }
    }

    # フォーム表示
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.Int32]$max_retries = 3
    for ([System.Int32]$i=0; $i -le $max_retries; $i++) {
        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # 入力値のチェック
            [System.String[]]$setting_parameters = @()
            $setting_parameters = @(
                $textbox_work_folder.Text,
                $textbox_resize_width.Text,
                $textbox_resize_height.Text
            )
            $messagecode = ValidateInputValues $setting_parameters

            # チェック結果が正常の場合
            if ($messagecode -eq [MESSAGECODE]::Successful) {
                $form = $null
                break
            }
        }
        else {
            $setting_parameters = @()
            $form = $null
            break
        }
        # 再試行回数を超過前の処理
        if ($i -eq $max_retries) {
            $messagecode = [MESSAGECODE]::Error_MaxRetries
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '再試行回数の超過'
            ShowMessagebox $messagebox_messages $messagebox_title
            $setting_parameters = @()
            $form = $null
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function SettingInputValues: setting_parameters [${setting_parameters}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $setting_parameters
}

#################################################################################
# 処理名　｜RetrieveMessage
# 機能　　｜メッセージ内容を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String（メッセージ内容）
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function RetrieveMessage {
    param (
        [MESSAGECODE]$target_code,
        [System.String]$append_message=''
    )
    [System.String]$return_messages = ''
    [System.String]$message = ''

    switch($target_code) {
        Successful                          {$message='正常終了';break}
        Abend                               {$message='異常終了';break}
        Cancel                              {$message='キャンセルしました。';break}
        Info_LoadedSettingfile              {$message='設定ファイルの読み込みが完了。';break}
        Info_WebpSkipBatchOcrKeywordcount   {$message='OCR一括処理でWebPファイルが対象となりました。WebPファイルはスキップします。';break}
        Info_ComplateConvertWebp            {$message='WebPファイルへの変換が正常終了しました。';break}
        Info_WebpSkipConvertWebp            {$message='WebP変換処理でWebPファイルが対象となりました。該当ファイルの処理はスキップし続行します。';break}
        Info_WebpSkipResizeImage            {$message='リサイズ処理でWebPファイルが対象となりました。該当ファイルの処理はスキップし続行します。';break}
        Info_WebpSkipSettingSizeIsBigger    {$message='設定により拡大処理は禁止となっています。（元画像のサイズ < 指定サイズ → 画像の拡大はしない）該当ファイルの処理はスキップし続行します。';break}
        Info_ComplateResizeImage            {$message='リサイズ処理が正常終了しました。';break}
        Confirm_ExecutionTool               {$message='ツールを実行します。';break}
        Confirm_OcrResult                   {$message='OCR結果は問題ないですか。';break}
        Confirm_ResizeImages                {$message='リサイズ処理を実行します。';break}
        Error_NotCore                       {$message='PowerShellエディションが「 Core 」ではありません。';break}
        Error_NotSupportedVersion           {$message='PowerShellバージョンがサポート対象外です。（バージョン7未満）';break}
        Error_NotWindows                    {$message='実行環境がWindows OSではありません。';break}
        Error_LoadingSettingfile            {$message='設定ファイルの読み込み処理でエラーが発生しました。';break}
        Error_NotExistsTargetpath           {$message='所定の場所に設定ファイルがありません。';break}
        Error_EmptyTargetfolder             {$message='作業フォルダーが空で指定されています。';break}
        Error_EmptyResizeValue              {$message='サイズ指定が空で指定されています。';break}
        Error_ZeroResizeValue               {$message='サイズ指定でゼロ（0）が指定されています。';break}
        Error_NotIntResizeValue             {$message='サイズ指定が正しい値で設定されていません。';break}
        Error_MaxRetries                    {$message='再試行回数を超過しました。';break}
        Error_EmptyOcrExepath               {$message='設定ファイル内に“OCR実行ファイルの場所”が空で指定されています。';break}
        Error_EmptyOcrTemppath              {$message='設定ファイル内に“OCRの一時ファイル作成場所”が空で指定されています。';break}
        Error_EmptyOcrSearchKeyword         {$message='設定ファイル内に“OCRの検索キーワード”が空で指定されています。';break}
        Error_NotExistsOcrExepath           {$message='設定ファイルにある“OCR実行ファイルの場所”にアクセスできません。';break}
        Error_NotExistsOcrTemppath          {$message='設定ファイルにある“OCRの一時ファイル作成場所”にアクセスできません。';break}
        Error_CopyTempfile                  {$message='OCR実行用の一時ファイルをコピー中にエラーが発生しました。';break}
        Error_ExecuteTesseractOcr           {$message='OCR実行中にエラーが発生しました。';break}
        Error_RemoveTempfile                {$message='OCR実行用の一時ファイルを削除中にエラーが発生しました。';break}
        Error_OverSizeForMonitor            {$message='モニターサイズよりも大きな画像に変換できません。';break}
        Error_ExecuteResize                 {$message='リサイズの実行中にエラーが発生しました。';break}
        Error_CreateResizeFolder            {$message='リトライ回数を超過した為、リサイズした画像ファイルを格納するフォルダー名を決定できませんでした。';break}
        Error_ResizefileSave                {$message='リサイズ後の保存処理でエラーが発生しました。';break}
        Error_ChangeExtension               {$message='拡張子の変更処理でエラーが発生しました。';break}
        Error_ConvertWebp                   {$message='WebPファイルの変換処理でエラーが発生しました。';break}
        Error_RemoveFile                    {$message='WebP変換後の元画像ファイルを削除する処理でエラーが発生しました。';break}
        default                             {break}
    }

    $sbtemp=New-Object System.Text.StringBuilder
    @("${message}`r`n",`
      "${append_message}`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $return_messages = $sbtemp.ToString()

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RetrieveMessage: return_messages [${return_messages}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $return_messages
}

#################################################################################
# 処理名　｜ShowMessagebox
# 機能　　｜メッセージボックスの表示
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function ShowMessagebox {
    param (
        [System.String]$messages,
        [System.String]$title,
        [System.String]$level='Information'
        # 指定可能なレベル一覧（$level）
        #   None
        #   Hand
        #   Error
        #   Stop
        #   Question
        #   Exclamation
        #   Waring
        #   Asterisk
        #   Information
    )

    [System.Windows.Forms.DialogResult]$dialog_result = [System.Windows.Forms.MessageBox]::Show($messages, $title, "OK", $level)
    
    switch($dialog_result) {
        {$_ -eq [System.Windows.Forms.DialogResult]::OK} {
            break
        }
    }
}
#################################################################################
# 処理名　｜RetrieveKeywordCount
# 機能　　｜キーワードを検索し件数を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String[]（検索キーワードとカウント数）
# 　　　　｜ - n次元目 項目01 検索キーワード
# 　　　　｜ - n次元目 項目02 キーワードの件数
# 引数　　｜targetfile   : 検索対象ファイル
# 　　　　｜keyword_lists: 検索キーワードリスト
# 　　　　｜casesensitive: 大文字・小文字を区別（true: 区別する、false: 区別しない）
#################################################################################
Function RetrieveKeywordCount {
    param (
        [System.String]$targetfile,
        [System.String[]]$keyword_lists,
        [System.Boolean]$casesensitive
    )
    [System.String]$keyword = ''
    [System.Int32]$keyword_count = 0
    [System.Object[]]$count_lists = @()

    # テキストファイル内の文字列チェック
    [System.String]$textdata = (Get-Content $targetfile)
    if ([string]::IsNullOrEmpty($textdata)){
        ### DEBUG ###
        if ($DEBUG_ON) {
            Write-Host '### DEBUG PRINT ###'
            Write-Host ''

            Write-Host "Function RetrieveKeywordCount: count_lists [${count_lists}]"

            Write-Host ''
            Write-Host '###################'
            Write-Host ''
            Write-Host ''
        }

        # 空のため早期リターン
        exit
    }

    # 複数キーワードで検索
    foreach($keyword in $keyword_lists) {
        if ($casesensitive) {
            # 大文字・小文字区別する
            $keyword_count = @(Select-String "${targetfile}" -Pattern "${keyword}" -CaseSensitive).Count
        }
        else {
            # 大文字・小文字区別しない
            $keyword_count = @(Select-String "${targetfile}" -Pattern "${keyword}").Count
        }

        $count_lists += ,@($keyword, $keyword_count)
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function RetrieveKeywordCount: count_lists [${count_lists}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $count_lists
}
#################################################################################
# 処理名　｜ShowCountlists
# 機能　　｜キーワード検索した結果を表示
#--------------------------------------------------------------------------------
# 戻り値　｜なし
# 引数　　｜count_lists: 検索キーワード毎に構成されたカウント数のリスト
# 　　　　｜targetfile : 検索対象ファイルのフルパス
#################################################################################
Function ShowCountlists {
    param (
        [System.Object[]]$count_lists,
        [System.String]$imagefile,
        [System.String]$targetfile
    )
    # リスト内の文字列チェック
    if ([string]::IsNullOrEmpty($count_lists)){
        ### DEBUG ###
        if ($DEBUG_ON) {
            Write-Host '### DEBUG PRINT ###'
            Write-Host ''

            Write-Host "Function ShowCountlists: count_lists [${count_lists}]"

            Write-Host ''
            Write-Host '###################'
            Write-Host ''
            Write-Host ''
        }

        # 空のため早期リターン
        exit
    }

    # 配列内で最大のバイト数（Shift-JIS）を取得
    [System.Object[]]$to_bytes = [Management.Automation.PSSerializer]::DeSerialize([Management.Automation.PSSerializer]::Serialize($count_lists))
    [System.Int32]$i = 0
    [System.Int32]$max_length = 0
    for ($i = 0; $i -lt $to_bytes.Count; $i++) {
        $to_bytes[$i][0] = [System.Text.Encoding]::GetEncoding("shift_jis").GetByteCount($to_bytes[$i][0])
        if ($max_length -lt $to_bytes[$i][0]) {
            $max_length = $to_bytes[$i][0]
        }
    }

    # 複数キーワードで検索
    Write-Host ' ============ 検索キーワード と 件数 ============ '
    Write-Host ''
    Write-Host " 元画像ファイル [${imagefile}]"
    Write-Host " 対象ファイル   [${targetfile}]"
    Write-Host ''
    Write-Host ' ------------------------------------------------ '
    Write-Host ''
    [System.Int32]$tab_count = 0
    [System.Int32]$tab_width = 4
    for ($i = 0; $i -lt $to_bytes.Count; $i++) {
        # 挿入するタブ数を計算
        $tab_count = [Math]::Ceiling(($max_length - [System.Int32]$to_bytes[$i][0]) / $tab_width)
        if ($tab_count -eq 0) {
            $tab_count = 1
        }

        if ($count_lists[$i][1] -eq 0) {
            Write-Host " 検索キーワード [$($count_lists[$i][0])]$("`t" * $tab_count)、件数 [$($count_lists[$i][1])件] "
        }
        else {
            Write-Host " 検索キーワード [$($count_lists[$i][0])]$("`t" * $tab_count)、件数 [$($count_lists[$i][1])件] " -ForegroundColor DarkRed
        }
    }
    Write-Host ''
    Write-Host ' ================================================ '
    Write-Host ''
    Write-Host ''
    Write-Host ''
}

#################################################################################
# 処理名　｜IsOnlyAsciiChar
# 機能　　｜ASCII文字だけで構成された文字列かチェック
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True; ASCII文字だけの文字列、False: ASCII文字列以外も含む文字列）
# 引数　　｜target_str   : 対象文字列
#################################################################################
Function IsOnlyAsciiChar {
    param (
        [System.String]$target_str
    )
    # すべての文字がASCII文字で構成されているか
    return ($target_str -match "^[\x00-\x7F]+$")
}

#################################################################################
# 処理名　｜GenerateFilename
# 機能　　｜使用するファイル名を生成
#--------------------------------------------------------------------------------
# 戻り値　｜String（名前を決めたファイル名。試行回数の超過もしくはエラーで作成できなかった場合はnullを返す）
# 引数　　｜target_dir      : 作業フォルダーのパス
# 　　　　｜filename        : ファイル名
# 　　　　｜max_retries     : 最大のリトライ回数
#################################################################################
Function GenerateFilename {
    param (
        [System.String]$target_dir,
        [System.String]$generate_filename,
        [System.Int32]$max_retries=30
    )
    [System.String]$newfilename = ''
    [System.String]$filename_without_ext = [System.IO.Path]::GetFileNameWithoutExtension(("${generate_filename}"))
    [System.String]$extension = ([System.IO.Path]::GetExtension("${generate_filename}")).ToLower()
    [System.Int32]$i = 0
    [System.String]$nowdate = (Get-Date).ToString("yyyyMMdd")
    [System.String]$number = ''
    for ($i=1; $i -le $max_retries; $i++) {
        # カウント数の数値を3桁で0埋めした文字列にする
        $number = "{0:000}" -f $i
        # 確認したいファイル名を生成
        $newfilename = "${filename_without_ext}_${nowdate}-${number}${extension}"
        # 作成したいフォルダー名の存在チェック
        if (-Not (Test-Path "${target_dir}\${newfilename}")) {
            break
        }

        # リトライ回数を超過し作成するフォルダー名を決定できなかった場合
        if ($i -eq $max_retries) {
            $newfilename = ''
        }
    }

    [System.String]$newfile_path = ''
    if ($newfilename -ne '') {
        $newfile_path = "${target_dir}\${newfilename}"
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function GenerateFilename: newfile_path [${newfile_path}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $newfile_path
}

#################################################################################
# 処理名　｜ExecuteTesseractOcr
# 機能　　｜OCRで画像の文字列をカウント
#--------------------------------------------------------------------------------
# 戻り値　｜Int[]（検索キーワードと言語毎のカウント数）
# 　　　　｜ - n次元目 項目01 検索キーワード n個目 + OCR英語の出力ファイル のカウント数
# 　　　　｜ - n次元目 項目02 検索キーワード n個目 + OCR日本語       （横書き）の出力ファイル のカウント数
# 　　　　｜ - n次元目 項目03 検索キーワード n個目 + OCR日本語 - vert（縦書き）の出力ファイル のカウント数
# 引数　　｜exepath         : OCR実行ファイルのフルパス
# 　　　　｜ocr_lang        : OCRの言語設定
# 　　　　｜argument_lists  : OCR実行時の引数
# 　　　　｜                    - 引数01 対象画像ファイルのフルパス 
# 　　　　｜                    - 引数02 出力するテキストファイル名（拡張子txtの記載なし）
# 　　　　｜                    - 引数03 OCRの言語設定（ocr_langを参照）
# 　　　　｜targetfile      : OCRで出力したテキストファイルのフルパス（拡張子txtを含む）
# 　　　　｜keyword_lists   : 検索するキーワードのリスト
# 　　　　｜casesensitive   : 検索する際に大文字・小文字を区別するか（True：区別する、False：区別しない）
#################################################################################
Function ExecuteTesseractOcr {
    param (
        [System.String]$exepath,
        [System.String]$ocr_lang,
        [System.String[]]$argument_lists,
        [System.String]$targetfile,
        [System.String[]]$keyword_lists,
        [System.Boolean]$casesensitive

    )
    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_message = ''

    [System.Object[]]$count_lists = @()
    [System.String]$imagefile = $argument_lists[0]

    # ファイル名にASCII文字以外が含まれている場合は、一時ファイルを作成
    [System.String]$original_fullpath = ''
    [System.String]$current_dir = ''
    [System.String]$temp_fullpath = ''
    [System.String]$filename_without_ext = [System.IO.Path]::GetFileNameWithoutExtension($argument_lists[0])
    [System.String]$extension = ([System.IO.Path]::GetExtension("$($argument_lists[0])")).ToLower()
    if (-Not(IsOnlyAsciiChar($filename_without_ext))) {
        $original_fullpath = $argument_lists[0]
        # 一時ファイルの作成場所を取得
        $current_dir = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($argument_lists[1]))
        $temp_fullpath = GenerateFilename $current_dir "OcrTempFile${extension}"

        try {
            Copy-Item $original_fullpath $temp_fullpath -Force
        }
        catch {
            $messagecode = [MESSAGECODE]::Error_CopyTempfile
            $messagecode_message = RetrieveMessage $messagecode
            Write-Host $messagecode_message -ForegroundColor DarkRed
        }
        # OCR実行時の引数に一時ファイルのパスを指定
        $argument_lists[0] = $temp_fullpath
    }

    # OCR実行
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        try {
            Start-Process -FilePath "${exepath}" -ArgumentList "${argument_lists}" -WindowStyle Hidden -Wait
            $count_lists = RetrieveKeywordCount "${targetfile}" $keyword_lists $casesensitive
    
            ShowCountlists $count_lists $imagefile $targetfile
        }
        catch {
            $messagecode = [MESSAGECODE]::Error_ExecuteTesseractOcr
            $messagecode_message = RetrieveMessage $messagecode
            Write-Host $messagecode_message -ForegroundColor DarkRed
        }
    }

    # 一時ファイルの削除
    if (($messagecode -eq [MESSAGECODE]::Successful) -and
        ($temp_fullpath -ne '')) {
        try {
            Remove-Item $temp_fullpath -Force
        }
        catch {
            $messagecode = [MESSAGECODE]::Error_RemoveTempfile
            $messagecode_message = RetrieveMessage $messagecode
            Write-Host $messagecode_message -ForegroundColor DarkRed
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ExecuteTesseractOcr: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}
#################################################################################
# 処理名　｜BatchOcrKeywordcount
# 機能　　｜OCRで画像の文字列をカウント
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜function_parameters: 設定ファイルの値
# 　　　　｜ - 項目01 作業フォルダー       : 画面での設定値 - 画像ファイルをチェックし変換する作業対象のフォルダー
# 　　　　｜ - 項目02 リサイズの横サイズ   : 画面での設定値 - リサイズ後の横サイズ（px）
# 　　　　｜ - 項目03 リサイズの縦サイズ   : 画面での設定値 - リサイズ後の縦サイズ（px）
# 　　　　｜ - 項目04 OCR実行ファイルのパス: 設定ファイルでの設定値 - Tesseract OCRのインストールパスの配下にあるEXEファイルまでのフルパス
# 　　　　｜ - 項目05 OCR結果の一時保存先　: 設定ファイルでの設定値 - Tesseract OCRのコマンド実行で出力するテキストデータの一時保存先
# 　　　　｜ - 項目06 検索するキーワード   : 設定ファイルでの設定値 - OCR結果のテキストデータを検索するキーワードを指定（複数指定可能）
# 　　　　｜ - 項目07 検索オプション       : 設定ファイルでの設定値 - OCR結果のテキストデータを検索する際に大文字・小文字を区別する（True：区別する、False：区別しない）
#################################################################################
Function BatchOcrKeywordcount {
    param (
        [System.Object[]]$function_parameters
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagebox_messages = ''
    [System.String]$messagebox_title = ''

    # 入力チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        #   OCR実行ファイルのパス
        if ($function_parameters[3] -eq '') {
            $messagecode = [MESSAGECODE]::Error_EmptyOcrExepath
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
        #   OCR結果の一時保存先
        if ($function_parameters[4] -eq '') {
            $messagecode = [MESSAGECODE]::Error_EmptyOcrTemppath
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
        #   検索するキーワード
        if ($function_parameters[5] -eq '') {
            $messagecode = [MESSAGECODE]::Error_EmptyOcrSearchKeyword
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '入力チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    # 存在チェック
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        #   OCR実行ファイルのパス
        if (-Not(Test-Path $function_parameters[3])) {
            $messagecode = [MESSAGECODE]::Error_NotExistsOcrExepath
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '存在チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
        #   OCR結果の一時保存先
        if (-Not(Test-Path $function_parameters[4])) {
            $messagecode = [MESSAGECODE]::Error_NotExistsOcrTemppath
            $messagebox_messages = RetrieveMessage $messagecode
            $messagebox_title = '存在チェック'
            ShowMessagebox $messagebox_messages $messagebox_title
        }
    }

    # 対象フォルダーにある対象ファイル毎にくり返し処理を開始
    [System.String]$item = ''
    [System.String]$item_basename = ''
    [System.String[]]$target_lists = Get-ChildItem -File "$($function_parameters[0])\*.*" -Include *.jpg,*jpeg,*.png,*.webp -Name
    [System.String]$ocr_lang = ''
    [System.String[]]$argument_lists = @()
    [System.String]$targetfile = ''
    [System.String]$exepath = $function_parameters[3]
    [System.String[]]$keyword_lists = $function_parameters[5].Split(',')
    [System.Boolean]$casesensitive = $function_parameters[6]
    [System.String]$extension = ''
    foreach($item in $target_lists) {
        # WebPの場合は処理をスキップ
        $extension = ([System.IO.Path]::GetExtension("${item}")).ToLower()
        if ($extension -eq '.webp') {
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
              "対象ファイル: [$($function_parameters[0])\$item]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_WebpSkipBatchOcrKeywordcount) $append_message
            Write-Host $prompt_message
            continue
        }
        $item_basename = (Get-Item "$($function_parameters[0])\$item").BaseName
        # OCR実行
        #   英語でOCRで文字列チェック
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            $ocr_lang = 'eng'
            $argument_lists = @("$($function_parameters[0])\$item", "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang)", '-l', $ocr_lang)
            $targetfile = "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang).txt"
            $messagecode = ExecuteTesseractOcr $exepath $ocr_lang $argument_lists $targetfile $keyword_lists $casesensitive
            if ($messagecode -ne [MESSAGECODE]::Successful) {
                exit
            }
        }

        #   日本語（横書き）でOCRで文字列チェック
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            $ocr_lang = 'jpn'
            $argument_lists = @("$($function_parameters[0])\$item", "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang)", '-l', $ocr_lang)
            $targetfile = "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang).txt"
            $messagecode = ExecuteTesseractOcr $exepath $ocr_lang $argument_lists $targetfile $keyword_lists $casesensitive
            if ($messagecode -ne [MESSAGECODE]::Successful) {
                exit
            }
        }

        #   日本語（縦書き）でOCRで文字列チェック
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            $ocr_lang = 'jpn_vert'
            $argument_lists = @("$($function_parameters[0])\$item", "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang)", '-l', $ocr_lang)
            $targetfile = "$($function_parameters[4])\Tesseract-OCR_$($ocr_lang).txt"
            $messagecode = ExecuteTesseractOcr $exepath $ocr_lang $argument_lists $targetfile $keyword_lists $casesensitive
            if ($messagecode -ne [MESSAGECODE]::Successful) {
                exit
            }
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function BatchOcrKeywordcount: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　｜CreateResizeFolder
# 機能　　｜リサイズした画像を格納するフォルダーを新規作成
#--------------------------------------------------------------------------------
# 戻り値　｜String（作成したフォルダー名。試行回数の超過もしくはエラーで作成できなかった場合は空文字を返す）
# 引数　　｜current_dir: 作業フォルダーのパス
# 　　　　｜foldername : 作成するフォルダー名
# 　　　　｜max_retries: 最大のリトライ回数
#################################################################################
Function CreateResizeFolder {
    param (
        [System.String]$current_dir,
        [System.String]$foldername,
        [System.Int32]$max_retries=30
    )
    [System.String]$newfoldername = $foldername
    [System.Int32]$i = 0
    [System.String]$nowdate = (Get-Date).ToString("yyyyMMdd")
    [System.String]$number = ''
    for ($i=1; $i -le $max_retries; $i++) {
        # カウント数の数値を3桁で0埋めした文字列にする
        $number = "{0:000}" -f $i
        # 作成したいフォルダー名を生成
        $newfoldername = "${foldername}_${nowdate}-${number}"
        # 作成したいフォルダー名の存在チェック
        if (-Not (Test-Path "${current_dir}\${newfoldername}")) {
            break
        }

        # リトライ回数を超過し作成するフォルダー名を決定できなかった場合
        if ($i -eq $max_retries) {
            $newfoldername = ''
        }
    }

    [System.String]$newfolder_path = ''
    if ($newfoldername -ne '') {
        $newfolder_path = "${current_dir}\${newfoldername}"
        try {
            New-Item -Path "${newfolder_path}" -Type Directory > $null
        }
        catch {
            $newfolder_path = ''
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function CreateResizeFolder: newfolder_path [${newfolder_path}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $newfolder_path
}

#################################################################################
# 処理名　｜ChangeExtension
# 機能　　｜指定したパスの拡張子を変換
#--------------------------------------------------------------------------------
# 戻り値　｜String（拡張子を変更したフルパス。拡張子が変わらない場合は空文字を返す）
# 引数　　｜targetpath   : 変換対象のフルパス
# 　　　　｜new_extension: 変換する拡張子
#################################################################################
function ChangeExtension {
    param (
        [System.String]$targetpath,
        [System.String]$new_extension
    )

    # 現在のパスからファイル名と拡張子を抽出
    [System.String]$filename_without_ext = [System.IO.Path]::GetFileNameWithoutExtension($targetpath)
    [System.String]$old_extension = [System.IO.Path]::GetExtension($targetpath)
    
    [System.String]$new_fullpath = ''
    $new_extension = $new_extension.ToLower()
    $old_extension = $old_extension.ToLower()
    # 新旧の拡張子が異なる場合のみ処理する
    if ($new_extension -ne $old_extension) {
        # 新しいフルパスを作成
        [System.String]$new_filename = $filename_without_ext + $new_extension
        $new_fullpath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($targetpath), $new_filename)
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ChangeExtension: new_fullpath [${new_fullpath}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $new_fullpath
}

#################################################################################
# 処理名　｜ConvertWebp
# 機能　　｜Webpへの変換
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜項目01 変換するcwebp.exeのフルパス
# 　　　　｜項目02 変換対象の画像ファイルのパス
#################################################################################
Function ConvertWebp {
    param (
        [System.String]$webp_exe_path,
        [System.String]$resize_output_path
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_message = ''

    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder
    [System.String]$prompt_message = ''
    [System.String]$append_message = ''

    # 拡張子のみ変更し変換後のフルパスの生成
    [System.String]$webp_output_path = ''
    try {
        $webp_output_path = ChangeExtension $resize_output_path '.webp'
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_ChangeExtension
        $messagecode_message = RetrieveMessage $messagecode
        Write-Host $messagecode_message -ForegroundColor DarkRed
    }

    # フルパスを生成できた場合
    if ($messagecode -eq [MESSAGECODE]::Successful) {
        if ($webp_output_path -ne '') {
            try {
                # WebPへの変換処理
                Start-Process -FilePath "${webp_exe_path}" -ArgumentList """${resize_output_path}"" -o ""${webp_output_path}""" -WindowStyle Hidden -Wait
            }
            catch {
                $messagecode = [MESSAGECODE]::Error_ConvertWebp
                $messagecode_message = RetrieveMessage $messagecode
                Write-Host $messagecode_message -ForegroundColor DarkRed
            }

            if ($messagecode -eq [MESSAGECODE]::Successful) {
                try {
                    # 変換できた後に元画像ファイルを削除
                    Remove-Item "${resize_output_path}" -Force
                }
                catch {
                    $messagecode = [MESSAGECODE]::Error_RemoveFile
                    $messagecode_message = RetrieveMessage $messagecode
                    Write-Host $messagecode_message -ForegroundColor DarkRed
                }
                
                # 通知
                $sbtemp=New-Object System.Text.StringBuilder
                @("`r`n",`
                "　変換したファイル: [${webp_output_path}]`r`n",`
                "　削除したファイル: [${resize_output_path}]`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $append_message = $sbtemp.ToString()
                $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_ComplateConvertWebp) $append_message
                Write-Host $prompt_message
            }
        }
        else {
            # 拡張子が既にWebPであるため通知して処理をスキップ
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
            "対象ファイル: [${resize_output_path}]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_WebpSkipConvertWebp) $append_message
            Write-Host $prompt_message
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ConvertWebp: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }

    return $messagecode
}

#################################################################################
# 処理名　｜ResizeImage
# 機能　　｜画像ファイルのリサイズ
#--------------------------------------------------------------------------------
# 戻り値　｜MESSAGECODE（enum）
# 引数　　｜setting_parameters[]
# 　　　　｜ - 項目01 作業フォルダー    
# 　　　　｜ - 項目02 リサイズの横サイズ
# 　　　　｜ - 項目03 リサイズの縦サイズ
# 　　　　｜ - 項目04 元画像のサイズよりも低いサイズで指定した場合にリサイズするか（True：リサイズする、False：リサイズしない）
# 　　　　｜ - 項目05 自動で作成する出力用のフォルダー名
# 　　　　｜ - 項目06 WEBPに変換するか（True：変換する、False：変換しない）
#################################################################################
Function ResizeImage {
    param (
        [System.Object[]]$function_parameters
    )

    [MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
    [System.String]$messagecode_message = ''

    # チェック処理（入力チェックや存在チェック）は、SettingInputValues で実施済みのため割愛。

    # リサイズ後の画像ファイルを格納するフォルダーを準備
    [System.String]$resize_output_path = CreateResizeFolder $function_parameters[0] $function_parameters[4]

    # リサイズ処理
    #   横サイズが指定された値を優先される。横サイズが指定なしで縦サイズのみ指定されている場合は、縦サイズを使用しリサイズ。
    [System.Int32]$before_width = 0
    [System.Int32]$before_height = 0
    [System.Int32]$after_width = 0
    [System.Int32]$after_height = 0
    [System.Double]$raito = 0
    [System.String]$item = ''
    [System.String[]]$target_lists = Get-ChildItem -File "$($function_parameters[0])\*.*" -Include *.jpg,*jpeg,*.png,*.webp -Name
    [System.Drawing.Bitmap]$before_image = $null
    [System.Drawing.Bitmap]$after_image = $null
    [System.Drawing.Graphics]$graphics = $null
    [System.String]$extension = ''
    foreach($item in $target_lists) {
        # WebPの場合は処理をスキップ
        $extension = ([System.IO.Path]::GetExtension("${item}")).ToLower()
        if ($extension -eq '.webp') {
            $sbtemp=New-Object System.Text.StringBuilder
            @("`r`n",`
              "対象ファイル: [$($function_parameters[0])\$item]`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $append_message = $sbtemp.ToString()
            $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_WebpSkipResizeImage) $append_message
            Write-Host $prompt_message
            continue
        }
        # 現在のサイズ取得
        $before_image = New-Object System.Drawing.Bitmap("$($function_parameters[0])\$item")
        $before_width  = $before_image.Width
        $before_height = $before_image.Height
        
        # リサイズする値を計算
        #   横サイズの値がある場合
        if ($function_parameters[1] -ne '') {
            # 横サイズでリサイズの比率を計算
            $after_width = $function_parameters[1]
            $raito = $after_width / $before_width
            # 縦サイズに反映
            $after_height = $before_height * $raito
        }
        #   縦サイズのみ値がある場合
        else {
            # 縦サイズでリサイズの比率を計算
            $after_height = $function_parameters[2]
            $raito = $after_height / $before_height
            # 横サイズに反映
            $after_width = $beforewitdh * $raito
        }

        # 指定した値が元画像のサイズ以上の場合に拡大しない場合
        if (-Not($function_parameters[3])) {
            # 横サイズの値がある場合
            if ($function_parameters[1] -ne '') {
                if ($after_width -ge $before_width) {
                    $sbtemp=New-Object System.Text.StringBuilder
                    @("`r`n",`
                      "対象ファイル: [$($function_parameters[0])\$item]`r`n")|
                    ForEach-Object{[void]$sbtemp.Append($_)}
                    $append_message = $sbtemp.ToString()
                    $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_WebpSkipSettingSizeIsBigger) $append_message
                    Write-Host $prompt_message
                    continue
                }
            }
            # 縦サイズのみ値がある場合
            else {
                if ($after_height -ge $before_height) {
                    $sbtemp=New-Object System.Text.StringBuilder
                    @("`r`n",`
                      "対象ファイル: [$($function_parameters[0])\$item]`r`n")|
                    ForEach-Object{[void]$sbtemp.Append($_)}
                    $append_message = $sbtemp.ToString()
                    $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_WebpSkipSettingSizeIsBigger) $append_message
                    Write-Host $prompt_message
                    continue
                }
            }
        }

        # リサイズの実行
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            try {
                $after_image = New-Object System.Drawing.Bitmap ($after_width, $after_height)
                $graphics = [System.Drawing.Graphics]::FromImage($after_image)
                $graphics.DrawImage($before_image, 0, 0, $after_width, $after_height)
            }
            catch {
                $messagecode = [MESSAGECODE]::Error_ExecuteResize
                $messagecode_message = RetrieveMessage $messagecode
                Write-Host $messagecode_message -ForegroundColor DarkRed
                break
            }
            finally {
                # リソース解放
                if ($null -ne $graphics) {
                    $graphics.Dispose()   
                }
                if ($null -ne $before_image) {
                    $before_image.Dispose()   
                }
            }
        }

        # フォルダー作成 と 保存
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            # 新規作成するフォルダー名を決定する
            if ($resize_output_path -eq '') {
                $messagecode = [MESSAGECODE]::Error_CreateResizeFolder
                $messagecode_message = RetrieveMessage $messagecode
                Write-Host $messagecode_message -ForegroundColor DarkRed
            }
            # 保存
            try {
                $after_image.Save("${resize_output_path}\${item}")
            }
            catch {
                $messagecode = [MESSAGECODE]::Error_ResizefileSave
                $messagecode_message = RetrieveMessage $messagecode
                Write-Host $messagecode_message -ForegroundColor DarkRed
                break
            }
            finally {
                # リソース解放
                if ($null -ne $after_image) {
                    $after_image.Dispose()   
                }
            }

            # 通知
            if ($messagecode -eq [MESSAGECODE]::Successful) {
                $sbtemp=New-Object System.Text.StringBuilder
                @("`r`n",`
                  "　リサイズしたファイル: [${resize_output_path}\${item}]`r`n")|
                ForEach-Object{[void]$sbtemp.Append($_)}
                $append_message = $sbtemp.ToString()
                $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_ComplateResizeImage) $append_message
                Write-Host $prompt_message
            }
        }

        # WEBP変換
        if ($messagecode -eq [MESSAGECODE]::Successful) {
            # 変換の有無
            if (($function_parameters[5]) -and
                ($resize_output_path -ne '')) {
                $messagecode = ConvertWebp $function_parameters[6] "${resize_output_path}\${item}"
                if ($messagecode -ne [MESSAGECODE]::Successful) {
                    break
                }
            }
        }
    }

    ### DEBUG ###
    if ($DEBUG_ON) {
        Write-Host '### DEBUG PRINT ###'
        Write-Host ''

        Write-Host "Function ResizeImage: messagecode [${messagecode}]"

        Write-Host ''
        Write-Host '###################'
        Write-Host ''
        Write-Host ''
    }
    
    return $messagecode
}
### Function <--- 終了 ---

### Main process --- 開始 --->
#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
# 初期設定
#   メッセージ関連
[MESSAGECODE]$messagecode = [MESSAGECODE]::Successful
[System.String]$prompt_title = ''
[System.String]$prompt_message = ''
[System.String]$messagecode_message = ''
[System.String]$append_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

#   PowerShell環境のチェック
$messagecode = VerificationExecutionEnv

#   設定ファイル読み込み
if ($messagecode -eq [MESSAGECODE]::Successful) {
    # ディレクトリの取得
    [System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
    Set-Location $current_dir'\..\..'
    [System.String]$root_dir = (Convert-Path .)

    # Configファイルのフルパスを作成  
    $sbtemp=New-Object System.Text.StringBuilder
    @("${current_dir}",`
      '\',`
      "${c_config_file}")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    [System.String]$config_fullpath = $sbtemp.ToString()

    # 読み込み処理
    try {
        [System.Collections.Hashtable]$config = (Get-Content $config_fullpath -Raw -Encoding UTF8).Replace('\','\\') | ConvertFrom-StringData

        # 変数に格納
        [System.String]$CONFIG_OCR_EXE_PATH=RemoveDoubleQuotes($config.ocr_exe_path)
        [System.String]$CONFIG_OCR_CHECK_KEYWORD=(RemoveDoubleQuotes($config.ocr_check_keyword))
        [System.Boolean]$CONFIG_OCR_CASE_SENSITIVE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.ocr_case_sensitive)))
        [System.Boolean]$CONFIG_AUTO_MODE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.auto_mode)))
        [System.String]$CONFIG_INPUT_FOLDER=RemoveDoubleQuotes($config.work_folder)
        [System.String]$CONFIG_RESIZE_WIDTH=RemoveDoubleQuotes($config.resize_width)
        [System.String]$CONFIG_RESIZE_HEIGHT=RemoveDoubleQuotes($config.resize_height)
        [System.Boolean]$CONFIG_SMALL_IMAGE_RESIZE=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.small_image_resize)))
        [System.String]$CONFIG_OUTPUT_FOLDERNAME=RemoveDoubleQuotes($config.output_foldername)
        [System.Boolean]$CONFIG_WEBP_CONVERT=[System.Convert]::ToBoolean((RemoveDoubleQuotes($config.webp_convert)))
        [System.String]$CONFIG_WEBP_EXE_PATH=RemoveDoubleQuotes($config.webp_exe_path)

        # 通知
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
        "対象ファイル: [${config_fullpath}]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $prompt_message = RetrieveMessage ([MESSAGECODE]::Info_LoadedSettingfile) $append_message
        Write-Host $prompt_message
    }
    catch {
        $messagecode = [MESSAGECODE]::Error_LoadingSettingfile
        $sbtemp=New-Object System.Text.StringBuilder
        @("`r`n",`
          "エラーの詳細: [${config_fullpath}$($_.Exception.Message)]`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $append_message = $sbtemp.ToString()
        $messagecode_message = RetrieveMessage ([MESSAGECODE]::Error_LoadingSettingfile) $append_message
    }   
}

#   入力値の設定
if ($messagecode -eq [MESSAGECODE]::Successful) {
    [System.Object[]]$function_parameters = @()
    [System.Object[]]$setting_parameters = @()
    if (-Not($CONFIG_AUTO_MODE)) {
        # 対話式の場合
        [System.String]$work_folder = $CONFIG_INPUT_FOLDER
        if ($work_folder -eq '') {
            $work_folder = $current_dir
        }
        $function_parameters = @(
            $root_dir,
            $work_folder,
            $CONFIG_RESIZE_WIDTH,
            $CONFIG_RESIZE_HEIGHT
        )

        # // TODO:  フォームから入力できる項目を増やした方が便利そう。
        #           現在、入力可能な項目
        #           ・作業用フォルダー
        #           ・リサイズする横サイズ
        #           ・リサイズする縦サイズ
        #
        #           環境依存する内容の為、追加すると便利そうな項目
        #           OCRの設定を追加
        #           ・Tesseract-OCR 実行ファイルの場所
        #           ・Tesseract-OCRの検索キーワード群
        #           ・Tesseract-OCRの検索キーワードのオプション（大文字・小文字を区別する）
        #           リサイズの設定を追加
        #           ・指定サイズより元画像が小さい場合の拡大有無（true：拡大する、false：拡大しない）
        #           ・WebP変換ツールのパス
        $setting_parameters = SettingInputValues $function_parameters
        if ($setting_parameters.Count -eq 0) {
            $messagecode = [MESSAGECODE]::Cancel
        }
    }
    else {
        # 自動実行の場合
        #   入力値のチェック
        [System.String[]]$setting_parameters = @()
        $setting_parameters = @(
            $CONFIG_INPUT_FOLDER
            $CONFIG_RESIZE_WIDTH
            $CONFIG_RESIZE_HEIGHT
        )
        $messagecode = ValidateInputValues $setting_parameters
    }
}

# // TODO:  「tesseract.exe」の存在チェック や バージョン確認 すると便利そう。
#           存在チェックで実行ファイルがない場合は、所定の場所にダウンロード。
#           バージョンチェックで最新ではない場合は、ダウンロードし所定の場所にある実行ファイルの置き換え。

#   ツール実行の有無確認
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $prompt_message = RetrieveMessage ([MESSAGECODE]::Confirm_ExecutionTool)
    If (ConfirmYesno $prompt_message) {
        # OCR実行
        $function_parameters = @(
            $setting_parameters[0],
            $setting_parameters[1],
            $setting_parameters[2],
            $CONFIG_OCR_EXE_PATH,
            # $CONFIG_OCR_TEMP_PATH,
            $setting_parameters[0],
            $CONFIG_OCR_CHECK_KEYWORD,
            $CONFIG_OCR_CASE_SENSITIVE
        )
        $messagecode = BatchOcrKeywordcount $function_parameters
    }
    else {
        $messagecode = [MESSAGECODE]::Cancel
    }
}

#   処理続行の確認
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $prompt_message = RetrieveMessage ([MESSAGECODE]::Confirm_OcrResult)
    $prompt_title = '処理続行の有無確認'
    If (-Not(ConfirmYesno $prompt_message)) {
        # Noの場合にキャンセル
        $messagecode = [MESSAGECODE]::Cancel
    }
}

#   リサイズ実行の有無確認
if ($messagecode -eq [MESSAGECODE]::Successful) {
    $prompt_message = RetrieveMessage ([MESSAGECODE]::Confirm_ResizeImages)
    $prompt_title = 'リサイズの有無確認'
    If (ConfirmYesno $prompt_message) {
        # リサイズ（有効な場合は合わせてWEBP変換）
        $function_parameters = @(
            $setting_parameters[0],
            $setting_parameters[1],
            $setting_parameters[2],
            $CONFIG_SMALL_IMAGE_RESIZE,
            $CONFIG_OUTPUT_FOLDERNAME,
            $CONFIG_WEBP_CONVERT,
            $CONFIG_WEBP_EXE_PATH
        )
        $messagecode = ResizeImage $function_parameters
    }
    else {
        $messagecode = [MESSAGECODE]::Cancel
    }
}

#   処理結果の表示
[System.String]$append_message = ''
$sbtemp=New-Object System.Text.StringBuilder
if ($messagecode -eq [MESSAGECODE]::Successful) {
    @("`r`n",`
      "メッセージコード: [${messagecode}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_message = RetrieveMessage $messagecode $append_message
    Write-Host $messagecode_message
}
else {
    @("`r`n",`
      "メッセージコード: [${messagecode}]`r`n",`
      $messagecode_message)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $append_message = $sbtemp.ToString()
    $messagecode_message = RetrieveMessage ([MESSAGECODE]::Abend) $append_message
    Write-Host $messagecode_message -ForegroundColor DarkRed
}

# 終了
exit $messagecode
### Main process <--- 終了 ---
