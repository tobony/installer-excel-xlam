# PowerShell script encoding setting
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 > $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUI 모드 설정
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)


# Get the executable path
$exePath = [System.IO.Path]::GetDirectoryName([System.Windows.Forms.Application]::ExecutablePath)
if ([string]::IsNullOrEmpty($exePath)) {
    # Fallback for PS1 script execution
    $exePath = $PSScriptRoot
    if ([string]::IsNullOrEmpty($exePath)) {
        # Final fallback to current directory
        $exePath = (Get-Location).Path
    }
}

# Source paths definition at the start
$SourcePath = Join-Path $exePath "src"  # default path
$InstallPath = "$env:APPDATA\Microsoft\AddIns"

# Write to log function
function Write-Log {
    param([string]$Message)
    $logTextBox.AppendText("$Message`r`n")
    $logTextBox.Select($logTextBox.Text.Length, 0)
    $logTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Excel Add-in Installer'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Log textbox
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10,130)
$logTextBox.Size = New-Object System.Drawing.Size(565,190)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = 'Vertical'
$logTextBox.ReadOnly = $true
$form.Controls.Add($logTextBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,330)
$progressBar.Size = New-Object System.Drawing.Size(565,20)
$form.Controls.Add($progressBar)

# Select folder button
$selectFolderButton = New-Object System.Windows.Forms.Button
$selectFolderButton.Location = New-Object System.Drawing.Point(10,70)
$selectFolderButton.Size = New-Object System.Drawing.Size(170,30)
$selectFolderButton.Text = 'Select source folder'
$form.Controls.Add($selectFolderButton)

# Install button (modified position)
$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(200,70)
$installButton.Size = New-Object System.Drawing.Size(170,30)
$installButton.Text = 'Install Excel Add-in'
$form.Controls.Add($installButton)

# Open Addin Folder button
$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Location = New-Object System.Drawing.Point(390,70)
$openFolderButton.Size = New-Object System.Drawing.Size(170,30)
$openFolderButton.Text = 'Open Addin Folder'
$form.Controls.Add($openFolderButton)

# Folder path label
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Location = New-Object System.Drawing.Point(10,105)
$folderLabel.Size = New-Object System.Drawing.Size(565,15)
$folderLabel.Text = "Source folder: $SourcePath"
$form.Controls.Add($folderLabel)

# Add folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select source folder containing Excel Add-in files"

# Select folder button click event
$selectFolderButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $script:SourcePath = $folderBrowser.SelectedPath
        $folderLabel.Text = "Source folder: $SourcePath"
        Write-Log "소스 폴더가 변경되었습니다: $SourcePath"

        # Verify source folder contains files
        $sourceFiles = Get-ChildItem -Path $SourcePath -File
        if ($sourceFiles.Count -eq 0) {
            Write-Log "[경고] 선택한 폴더에 파일이 없습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "선택한 폴더에 설치할 파일이 없습니다.`n다른 폴더를 선택해주세요.",
                "경고",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        } else {
            Write-Log "→ 발견된 파일 수: $($sourceFiles.Count)"
        }
    }
})

# Open Addin Folder button click event
$openFolderButton.Add_Click({
    if (Test-Path $InstallPath) {
        Start-Process "explorer.exe" -ArgumentList $InstallPath
        Write-Log "Add-in 폴더를 열었습니다: $InstallPath"
    } else {
        Write-Log "[경고] Add-in 폴더를 찾을 수 없습니다: $InstallPath"
        [System.Windows.Forms.MessageBox]::Show(
            "Add-in 폴더를 찾을 수 없습니다.`n경로: $InstallPath",
            "경고",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Description label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(565,40)
$label.Text = 'This tool will install Excel Add-in from the src folder and register it in Excel.'
$form.Controls.Add($label)

# Form load event handler
$form.Add_Shown({
    # 폼이 완전히 로드된 후 Excel 프로세스 체크
    $installButton.Enabled = $false
    Write-Log "Excel 프로세스 확인 중..."
    
    $excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Excel이 실행 중입니다. 설치를 진행하려면 Excel을 종료해야 합니다.`n`n계속 진행하시겠습니까?",
            "Excel 실행 확인",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            Write-Log "사용자가 Excel 종료를 취소했습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "Excel을 종료한 후 설치 프로그램을 다시 실행해주세요.",
                "설치 취소",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            $form.Close()
            return
        }

        Write-Log "Excel 프로세스 종료 중..."
        foreach ($proc in $excelProcesses) {
            $proc.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (-not $proc.HasExited) {
                $proc.Kill()
            }
        }

        Start-Sleep -Seconds 2
        $remainingExcel = Get-Process excel -ErrorAction SilentlyContinue
        if ($remainingExcel) {
            Write-Log "Excel을 완전히 종료하지 못했습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "Excel을 완전히 종료하지 못했습니다.`n모든 Excel 작업을 저장하고 직접 종료한 후 다시 시도해주세요.",
                "Excel 종료 실패",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            $form.Close()
            return
        }
        Write-Log "Excel 프로세스가 종료되었습니다. 버튼을 클릭해서 다음 절차를 진행하세요."
    } else {
        Write-Log "Excel이 실행중이지 않습니다. 버튼을 클릭해서 다음 절차를 진행하세요."
    }
    
    $installButton.Enabled = $true
})

# Installation function
function Install-AddIn {
    $installButton.Enabled = $false
    $progressBar.Value = 0
    
    try {
        Write-Log "설치를 시작합니다..."
        $progressBar.Value = 10
        
        # Check source folder
        if (-not (Test-Path $SourcePath)) {
            Write-Log "[오류] 소스 폴더를 찾을 수 없습니다: $SourcePath"
            [System.Windows.Forms.MessageBox]::Show(
                "소스 폴더를 찾을 수 없습니다.`n폴더 경로를 확인해주세요.`n경로: $SourcePath",
                "오류",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Unblock files in source directory
        Write-Log "파일 차단 해제 중..."
        $sourceFiles = Get-ChildItem -Path $SourcePath -File
        foreach ($file in $sourceFiles) {
            try {
                Unblock-File -Path $file.FullName -ErrorAction Stop
                Write-Log "→ 파일 차단 해제 완료: $($file.Name)"
            }
            catch {
                Write-Log "[경고] 파일 차단 해제 실패 ($($file.Name)): $_"
            }
        }
        
        # Continue with the rest of the installation process...
        $progressBar.Value = 30
        
        # Check source files
        if ($sourceFiles.Count -eq 0) {
            Write-Log "[오류] src 폴더가 비어있습니다."
            [System.Windows.Forms.MessageBox]::Show(
                "src 폴더에 설치할 파일이 없습니다.",
                "오류",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $progressBar.Value = 40
        # # Create install directory if not exists
        # if (-not (Test-Path $InstallPath)) {
        #     New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        #     Write-Log "AddIn 폴더 생성됨: $InstallPath"
        # }
        # Validate installation path
        Write-Log "AddIn 설치 경로 확인 중: $InstallPath"
        if ([string]::IsNullOrEmpty($InstallPath)) {
            throw "Installation path is null or empty"
        }
        
        # Create or verify install directory
        try {
            if (-not (Test-Path $InstallPath)) {
                New-Item -ItemType Directory -Path $InstallPath -Force -ErrorAction Stop | Out-Null
                Write-Log "AddIn 폴더 생성됨: $InstallPath"
            } else {
                Write-Log "기존 AddIn 폴더 사용: $InstallPath"
                # Verify write permissions
                $testFile = Join-Path $InstallPath "test.tmp"
                try {
                    [System.IO.File]::WriteAllText($testFile, "")
                    Remove-Item $testFile -Force
                } catch {
                    throw "AddIn 폴더에 쓰기 권한이 없습니다: $InstallPath"
                }
            }
        } catch {
            Write-Log "[오류] AddIn 폴더 접근 실패: $_"
            throw "Failed to create or access AddIn directory: $_"
        }
        
        $progressBar.Value = 50
        # # Registry path for Excel add-ins
        # $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
        # $maxOpenKey = 0
        
        # # Check existing OPEN keys
        # if (Test-Path $registryPath) {
        #     $openKeys = Get-ItemProperty -Path $registryPath -Name 'OPEN*' -ErrorAction SilentlyContinue
        #     if ($openKeys) {
        #         $openKeys.PSObject.Properties | Where-Object { $_.Name -like 'OPEN*' } | ForEach-Object {
        #             $keyNum = [int]($_.Name -replace 'OPEN', '')
        #             if ($keyNum -gt $maxOpenKey) {
        #                 $maxOpenKey = $keyNum
        #             }
        #         }
        #     }
        # }

        # Registry path for Excel add-ins
        Write-Log "Excel 버전 확인 중..."
        try {
            # Determine Excel version and set registry path accordingly
            $excelVersion = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\' -ErrorAction Stop).PSObject.Properties.Name | 
                Where-Object { $_ -like '16.*' -or $_ -like '15.*' } | 
                Sort-Object -Descending | 
                Select-Object -First 1

            if ([string]::IsNullOrEmpty($excelVersion)) {
                Write-Log "Excel 버전을 찾을 수 없습니다. Office 365 기본값 사용."
                $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' # Default for 365
            } else {
                Write-Log "감지된 Excel 버전: $excelVersion"
                switch ($excelVersion) {
                    '16.0' { $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' }
                    '15.0' { $registryPath = 'HKCU:\Software\Microsoft\Office\15.0\Excel\Options' }
                    default { $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' }
                }
            }
        } catch {
            Write-Log "[경고] Excel 버전 확인 실패. Office 365 기본값 사용."
            $registryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options' # Fallback for 365
        }
        Write-Log "레지스트리 경로: $registryPath"
        $maxOpenKey = 0

        # Check existing OPEN keys
        Write-Log "레지스트리 키 확인 중..."
        try {
            if (Test-Path $registryPath) {
                $openKeys = Get-ItemProperty -Path $registryPath -Name 'OPEN*' -ErrorAction SilentlyContinue
                if ($openKeys) {
                    $openKeys.PSObject.Properties | Where-Object { $_.Name -like 'OPEN*' } | ForEach-Object {
                        if ($_.Name -match 'OPEN(\d+)') {
                            $keyNum = [int]$matches[1]
                            if ($keyNum -gt $maxOpenKey) {
                                $maxOpenKey = $keyNum
                            }
                        }
                    }
                    Write-Log "기존 OPEN 키 최대값: $maxOpenKey"
                } else {
                    Write-Log "기존 OPEN 키가 없습니다."
                }
            } else {
                Write-Log "레지스트리 경로가 없습니다. 새로 생성됩니다."
                New-Item -Path $registryPath -Force | Out-Null
            }
        } catch {
            Write-Log "[경고] 레지스트리 키 확인 중 오류: $_"
            # Continue with maxOpenKey = 0
        }

        
        $progressBar.Value = 60
        # Process each file
        $fileCount = $sourceFiles.Count
        $current = 0
        
        foreach ($file in $sourceFiles) {
            $current++
            $progressValue = 60 + (30 * ($current / $fileCount))
            $progressBar.Value = [math]::Min(90, $progressValue)
            
            $destPath = Join-Path $InstallPath $file.Name
            try {
                # Copy file
                Copy-Item -Path $file.FullName -Destination $destPath -Force
                Write-Log "→ 파일 복사 완료: $($file.Name)"
                
                # Register if it's an XLAM file
                if ($file.Extension -eq '.xlam') {
                    try {
                        $maxOpenKey++
                        $newKeyName = "OPEN$maxOpenKey"
                        
                        if (-not (Test-Path $registryPath)) {
                            Write-Log "레지스트리 경로 생성 중: $registryPath"
                            New-Item -Path $registryPath -Force -ErrorAction Stop | Out-Null
                        }
                        
                        Write-Log "Add-in 등록 중: $($file.Name)"
                        New-ItemProperty -Path $registryPath -Name $newKeyName -Value $destPath -PropertyType String -Force -ErrorAction Stop | Out-Null
                        Write-Log "→ Add-in 등록 완료: $($file.Name) (키: $newKeyName)"
                    } catch {
                        Write-Log "[오류] Add-in 등록 실패: $_"
                        throw "Failed to register Add-in in registry: $_"
                    }
                }
            }
            catch {
                Write-Log "[오류] 파일 처리 실패 ($($file.Name)): $_"
                [System.Windows.Forms.MessageBox]::Show(
                    "파일 처리 중 오류가 발생했습니다: $($file.Name)`n$_",
                    "오류",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        
        $progressBar.Value = 100
        Write-Log "`n[완료] 설치가 완료되었습니다."
        Write-Log "Excel을 시작하면 Add-in을 사용할 수 있습니다."
        
        [System.Windows.Forms.MessageBox]::Show(
            "설치가 완료되었습니다.`nExcel을 시작하면 Add-in을 사용할 수 있습니다.",
            "설치 완료",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Write-Log "[오류] 설치 중 오류가 발생했습니다: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "설치 중 오류가 발생했습니다:`n$_",
            "오류",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $installButton.Enabled = $true
        $progressBar.Value = 0
    }
}

# Install button click event
$installButton.Add_Click({ Install-AddIn })

# Check and request admin privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "이 프로그램은 관리자 권한이 필요합니다. 관리자 권한으로 다시 시작하시겠습니까?",
        "관리자 권한 필요",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processStartInfo.FileName = "powershell.exe"
        $processStartInfo.Arguments = "-File `"$($myinvocation.mycommand.definition)`""
        $processStartInfo.Verb = "runas"
        [System.Diagnostics.Process]::Start($processStartInfo)
    }
    exit
}

[System.Windows.Forms.Application]::Run($form)