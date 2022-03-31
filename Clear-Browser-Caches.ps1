# -*- coding: japanese-cp932-dos; mode: powershell; -*-

# ユーザーのダウンロードファイルを何日前まで残すか(0日の場合はダウンロードファイルを削除しない)
$DayOfUserDownloadFileToKeep = 0

# デバッグモードの場合は $true に
$SetWhatIf = $false

# ユーザー名を指定すると他のユーザーに対しては実行されない
# $null を指定した場合には Cドライブ直下の全てのユーザーが対象
$ForOneUser = $null

#
# ====================================================================
#   システム関係
# ====================================================================
# 管理者権限を必須とするか
$NeedAdminRole = $true

if ($NeedAdminRole) {
    function CheckAdmin {
        $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
        $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if ((CheckAdmin) -eq $false) {
        Write-Host -ForegroundColor Red "------------------------------------------------"
        Write-Host -ForegroundColor Red "  このスクリプトは管理者権限で実行してください  "
        Write-Host -ForegroundColor Red "------------------------------------------------"
        Exit
    }

}

# ====================================================================
#   Class
# ====================================================================

class Logger {
    [string] $logFile

    Logger ($MyInvocation) {
        $this.logFile = Join-Path (Get-Location) ((Get-Date -Format "yyyy_MMd_HHmm") + ".log")
    }

    [void]
    WriteLog($contents) {
        foreach($line in $contents.split("`n")) {
            Write-Output $line | Out-File -FilePath $this.logFile -Encoding Default -append
        }
    }
}

class AppMessage
{
    hidden [String] $title
    hidden [Logger] $logger

    AppMessage($logger) {
        $this.title = $null
        $this.logger = $logger
    }

    [void]
    SetTitle([string]$v) {
        $this.title = $v
    }

    [void]
    Info([string]$message) {
        $contents  = $this.GetMessageWithTitle($message)
        $this.WriteOut($contents, "Green")
    }

    [void]
    Skip([string]$message) {
        $contents  = $this.GetMessageWithTitle($message)
        $this.WriteOut($contents, "DarkGray")
    }

    [void]
    Done([string]$message) {
        $contents  = $this.GetMessageWithTitle($message)
        $this.WriteOut($contents, "DarkGray")
    }

    [void]
    WhatIf([string]$message) {
        $contents  = $this.GetMessageWithTitle($message)
        $this.WriteOut($contents, "Cyan")
    }

    [void]
    Warning([string]$message) {
        $contents  = $this.GetMessageWithTitle($message)
        $this.WriteOut($contents, "Red")
    }


    [void]
    Head1([string]$message) {
        $contents  = "======================================================================`n"
        $contents += "  " + $this.GetMessageWithTitle($message) + "`n"
        $contents += "======================================================================"
        $this.WriteOut($contents, "Yellow")
    }

    [void]
    Head2([string]$message) {
        $contents  = "`n----- " + $this.GetMessageWithTitle($message) + " -----"
        $this.WriteOut($contents, "Yellow")
    }

    [string]
    GetMessageWithTitle([string]$message) {
        $date = Get-Date -Format "[yyyy-MM-d HH:mm]"
        if ($this.title) {
            $message = "$date[" + $this.title + "] " + $message
        }
        return $message
    }

    [void]
    WriteOut([string]$contents, [string]$color = "") {
        Write-Host -ForegroundColor $color $contents
        $this.logger.WriteLog($contents)
    }
}

class TargetUser
{
    hidden [String] $name
    hidden [Bool] $is_whatif
    hidden [AppMessage] $message

    TargetUser($message) {
        $this.message = $message
    }

    [string]
    GetName() {
        return $this.name
    }

    [void]
    SetName([string]$name) {
        $this.name = $name
        $this.message.SetTitle($name)
        $this.message.Head1("処理を開始します")
    }

    [bool]
    GetWhatIf() {
        return $this.is_whatif
    }

    [string]
    GetLocalAppDataPath() {
        return "C:\Users\" + $this.GetName() + "\AppData\Local"
    }

    [string]
    GetDownloadsPath() {
        return "C:\Users\" + $this.GetName() + "\Downloads"
    }

    [void]
    SetWhatIf([Bool]$v) {
        $this.is_whatif = $v
    }

    [void]
    RemoveItemRecusive([string]$path) {
        if(Test-Path $path) {
            if ($this.GetWhatIf()) {
                $this.message.WhatIf("Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue -Verbose")
            } else {
                $this.message.Info($path + " を削除します")
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue -Verbose
            }
            $this.message.Done("$path を削除しました")
        } else {
            $this.message.Skip("$path は存在しません。スキップします。")
        }
    }

    [void]
    ClearCachesGoogleChrome() {
        $app_base_path = Join-Path $this.GetLocalAppDataPath() "Google\Chrome\User Data"
        if($this.AppBasePathExist("Google Chrome", $app_base_path)) {
            # Clear Default Caches
            $this.ClearCachesGoogleChromeProfile($app_base_path, "Default")
            # Clear Other Profile Caches
            $profiles = Get-ChildItem -Path $app_base_path | Select-Object Name | Where-Object Name -Like "Profile*"
            foreach ($profile_ in $profiles) {
                $profile_name = $profile_.Name
                $this.ClearCachesGoogleChromeProfile($app_base_path, $profile_name)
            }
        }
    }

    [void]
    ClearCachesVivaldi() {
        $app_base_path = Join-Path $this.GetLocalAppDataPath() "Vivaldi\User Data"
        if($this.AppBasePathExist("Vivaldi", $app_base_path)) {
            # Clear Default Caches
            $this.ClearCachesGoogleChromeProfile($app_base_path, "Default")
            # Clear Other Profile Caches
            $profiles = Get-ChildItem -Path $app_base_path | Select-Object Name | Where-Object Name -Like "Profile*"
            foreach ($profile_ in $profiles) {
                $profile_name = $profile_.Name
                $this.ClearCachesGoogleChromeProfile($app_base_path, $profile_name)
            }
        }
    }

    [void]
    ClearCachesGoogleChromeProfile([string]$app_base_path, $profile_name) {
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Code Cache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Service Worker\CacheStorage\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Service Worker\ScriptCache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cache2\entries\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cookies"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Media Cache"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cookies-Journal"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\JumpListIconsOld"))
    }

    [void]
    ClearCachesInternetExplorer() {
        $app_base_path = Join-Path $this.GetLocalAppDataPath() "Microsoft\Windows"
        if($this.AppBasePathExist("Internet Explorer", $app_base_path)) {
            $this.RemoveItemRecusive((Join-Path $app_base_path "Temporary Internet Files\*"))
            $this.RemoveItemRecusive((Join-Path $app_base_path "INetCache\*"))
            $this.RemoveItemRecusive((Join-Path $app_base_path "WebCache\*"))
        }
    }

    [void]
    ClearCachesEdgeChronium() {
        $app_base_path = Join-Path $this.GetLocalAppDataPath() "Microsoft\Edge\User Data"
        if($this.AppBasePathExist("Edge(Chronium)", $app_base_path)) {
            # Clear Default Caches
            $this.ClearCachesEdgeChroniumProfile($app_base_path, "Default")
            # Clear Other Profile Caches
            $profiles = Get-ChildItem -Path $app_base_path | Select-Object Name | Where-Object Name -Like "Profile*"
            foreach ($profile_ in $profiles) {
                $profile_name = $profile_.Name
                $this.ClearCachesEdgeChroniumProfile($app_base_path, $profile_name)
            }
        }
    }

    [void]
    ClearCachesEdgeChroniumProfile([string]$app_base_path, $profile_name) {
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Code Cache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Service Worker\CacheStorage\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Service Worker\ScriptCache\*"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cookies"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\Cookies-Journal"))
        $this.RemoveItemRecusive((Join-Path $app_base_path "$profile_name\EdgeDWriteFontCache"))
    }


    [void]
    ClearUserTempCaches() {
        $this.RemoveItemRecusive((Join-Path $this.GetLocalAppDataPath() "Temp\*"))
        $this.RemoveItemRecusive((Join-Path $this.GetLocalAppDataPath() "Microsoft\Windows\WER\*"))
        $this.RemoveItemRecusive((Join-Path $this.GetLocalAppDataPath() "Microsoft\Windows\AppCache\*"))
        $this.RemoveItemRecusive((Join-Path $this.GetLocalAppDataPath() "CrashDumps\*"))
    }

    [void]
    DeleteOldDownloads([int]$day_of_life = 0) {
        if ($day_of_life -eq 0) {
            $this.message.Skip("ダウンロードファイルは削除しません")
        } else {
            $this.message.Head2("ダウンロードファイルを削除します")
            $delete_until_date = (Get-Date).AddDays(-1 * $day_of_life)
            $old_files = Get-ChildItem -Path $this.GetDownloadsPath() -Recurse -File -ErrorAction SilentlyContinue `
              | Where-Object LastWriteTime -LT $delete_until_date
            foreach ($file in $old_files) {
                if ($this.GetWhatIf()) {
                    $this.message.WhatIf("Remove-Item -Path """ + $file.FullName + """-Force -ErrorAction SilentlyContinue -Verbose")
                } else {
                    $this.message.Info($file.FullName + " を削除します")
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue -Verbose
                }
            }
        }
    }

    [bool]
    AppBasePathExist([string]$app_name, [string]$app_base_path) {
        if(Test-Path $app_base_path) {
            $this.message.Head2("$app_name のキャッシュ等をクリアします")
            return $true
        } else {
            $this.message.Skip("$app_name の基本パス $app_base_path が存在しません。スキップします。")
            return $false
        }
    }

}


# ====================================================================
#   Functions
# ====================================================================
function clear_user_caches([string]$username, $message) {
    $user = [TargetUser]::new($message)
    $user.SetWhatIf($SetWhatIf)
    $user.SetName($username)
    $user.ClearCachesInternetExplorer()
    $user.ClearCachesGoogleChrome()
    $user.ClearCachesVivaldi()
    $user.ClearCachesEdgeChronium()
    $user.DeleteOldDownloads($DayOfUserDownloadFileToKeep)
}

# ====================================================================
#   Main
# ====================================================================

$logger = [Logger]::new($MyInvocation)

[void][reflection.assembly]::LoadWithPartialName("System.DirectoryServices")
[void][reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")

$context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine)
$message = [AppMessage]::new($logger)

if ($ForOneUser) {
    clear_user_caches $ForOneUser $message
} else {
    $users = Get-ChildItem "C:\Users" | Select-Object Name
    Foreach($user in $users) {
        $user_principal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($context, [System.DirectoryServices.AccountManagement.IdentityType]::Name, $user.Name)
        if($user_principal) {
            $message.Head1($user.Name + "はローカルユーザーです。スキップします。")
        } else {
            clear_user_caches $user.Name $message
        }
    }
}
