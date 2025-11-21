; =================================================================
; ahkURLforCitavi.ahk (AutoHotkey v2.0 版本)
; 功能: 通过 ahk://citavi/goto?type=Ref&id=xxx&project=xxx&projectType=xxx URL跳转到Citavi
; 用途: 由Windows注册表协议调用
; =================================================================

; --- 全局错误处理 ---
MyErrHandler(*) {
    errMsg := "脚本发生错误:`n`n"
    errMsg .= "错误信息: " . A_LastError . "`n"
    errMsg .= "错误行号: " . A_LineNumber . "`n"
    errMsg .= "所在文件: " . A_ScriptFullPath . "`n"
    
    A_Clipboard := errMsg
    MsgBox("发生错误，错误详情已复制到剪贴板。`n`n" . errMsg)
}
OnError(MyErrHandler)

#SingleInstance Force

; =================================================================
; 辅助函数定义区域
; =================================================================

/**
 * 在指定时间内智能等待并激活目标窗口
 * @param winTitle 要等待和激活的窗口标题
 * @param maxWaitTimeSec 最大等待时间（秒），默认为25
 * @returns {boolean} 成功返回 true，超时返回 false
 */
WaitForAndActivateWindow(winTitle, maxWaitTimeSec := 60) {
    CheckInterval := 2 ; 检查间隔（秒）
    MaxLoops := maxWaitTimeSec / CheckInterval
    CurrentLoop := 0

    Loop {
        if WinExist(winTitle) {
            WinActivate() ; 激活 WinExist 刚刚找到的窗口
            return true ; 成功
        }
        
        CurrentLoop++
        if (CurrentLoop >= MaxLoops)
            break ; 超时
        
        Sleep(CheckInterval * 1000)
    }
    return false ; 失败
}

/**
 * 显示一个TrayTip，并在指定秒数后隐藏
 * @param message 要显示的消息文本
 * @param duration 显示时长（秒），默认为5
 */
ShowTimedTrayTip(message, duration := 5) {
    TrayTip(message)
    Sleep(duration * 1000)
    HideTrayTip()
}

HideTrayTip() {
    TrayTip  ; 尝试以普通的方式隐藏它.
    if SubStr(A_OSVersion,1,3) = "10." {
        A_IconHidden := true
        Sleep 200  ; 可能有必要调整 sleep 的时间.
        A_IconHidden := false
    }
}

/**
 * 切换输入法
 * @param dwLayout 输入法代码
 */
SwitchIME(dwLayout) {
    HKL := DllCall("LoadKeyboardLayout", "Str", dwLayout, "UInt", 1)
    SendMessage(0x50, 0, HKL, , "A")
}

; ==============================================
; 中文URL解码函数 (您提供的优秀方案)
; ==============================================
ChineseUrlDecode(url) {
    try {
        static doc := ComObject("htmlfile")
        static js := doc.parentWindow
        if !doc.documentElement {
            doc.write('<meta http-equiv="X-UA-Compatible" content="IE=9">')
            doc.close()
        }
        return js.decodeURIComponent(url)
    } catch {
        try {
            local decodedUrl := url
            local result := DllCall(
                'Shlwapi.dll\UrlUnescape',
                'Str', decodedUrl,
                'Ptr', 0,
                'UInt', 0,
                'UInt', 0x00100000 | 0x00040000,
                'UInt'
            )
            if result = 0 {
                return decodedUrl
            } else {
                MsgBox("URL解码失败，返回原始字符串")
                return url
            }
        } catch {
            MsgBox("所有解码方法都失败: ")
            return url
        }
    }
}


; =================================================================
; 主入口逻辑
; =================================================================
if (A_Args.Length > 0) {
    UrlToProcess := A_Args[1]
    
    ; 使用正则表达式匹配并提取 type, id, project 和 projectType 参数
    if (RegExMatch(UrlToProcess, "ahk://citavi/goto\?type=([a-zA-Z]+)&id=([a-f0-9-]{36})(?:&project=([^&]+))?(?:&projectType=([^&]+))?", &Match)) {
        IdType := Match[1]
        IdValue := Match[2]
        ProjectName := Match[3] ; 获取项目名/路径
        ProjectType := Match[4] ; 获取项目类型

        ; 如果没有project参数，则无法继续
        if (!ProjectName) {
            MsgBox("错误：URL中缺少必要的 project 参数。")
            ExitApp
        }

        ; 如果没有projectType参数，则默认为DesktopSQLite
        if (!ProjectType) {
            ProjectType := "DesktopSQLite"
        }

        ; 对项目名/路径进行URL解码
		ProjectName := ChineseUrlDecode(ProjectName)
        ; --- 智能查找并激活目标项目的主窗口 ---
        TargetFound := false
        FullWinTitle := "" ; 初始化

        ; --- 根据项目类型执行不同逻辑 ---
        if (ProjectType = "DesktopSQLite") {
            ; --- 逻辑1: 处理本地SQLite项目 ---
            ; 1. 从项目路径中提取纯文件名（不含扩展名）
			
			ProjectNameWin := StrReplace(ProjectName, "/", "\")
            SplitPath(ProjectNameWin, , , , &OutNameNoExt)
            ProjFileNameNoExt := OutNameNoExt
            ; 2. 循环检查三种可能的窗口标题
            Workspaces := ["Reference Editor", "Knowledge Organizer", "Task Planner"]
            for index, workspaceName in Workspaces {
                TargetTitlePattern := ProjFileNameNoExt . ": " . workspaceName
                FullWinTitle := TargetTitlePattern . " ahk_exe Citavi.exe"
                
                if WinExist(FullWinTitle) {
                    try {
                        WinActivate(FullWinTitle)
                        TargetFound := true
                        break ; 找到后立即退出循环
                    }
                }
            }

            ; 如果没找到目标窗口，则尝试打开项目
            if (!TargetFound) {
                ; 检查Citavi进程是否存在
                if not WinExist("ahk_exe Citavi.exe") {
                    ; Citavi完全没运行，直接打开项目
                    Run(ProjectName, , "Wait")
                    ShowTimedTrayTip("Citavi未运行，正在打开项目: " . ProjectName)
                    
                    if !WaitForAndActivateWindow(FullWinTitle, 60) {
                        MsgBox("Citavi启动或窗口激活超时，在60秒内未找到目标窗口。")
                        ExitApp
                    }
                } else {
                    ; Citavi在运行，但目标项目窗口没找到（可能项目没打开）
                    Run(ProjectName, , "Wait")
                    ShowTimedTrayTip("Citavi已运行，正在尝试打开项目: " . ProjectName)
                    if !WaitForAndActivateWindow(FullWinTitle, 60) {
                        MsgBox("项目打开或窗口激活超时，在60秒内未找到目标窗口。")
                        ExitApp
                    }
                }
            }
        } else {
            ; --- 逻辑2: 处理服务器或云端项目 ---
            ; ProjectName是纯项目名，直接用它构建窗口标题
            Workspaces := ["Reference Editor", "Knowledge Organizer", "Task Planner"]
            for index, workspaceName in Workspaces {
                TargetTitlePattern := ProjectName . ": " . workspaceName
                FullWinTitle := TargetTitlePattern . " ahk_exe Citavi.exe"

                if WinExist(FullWinTitle) {
                    try {
                        WinActivate(FullWinTitle)
                        TargetFound := true
                        break
                    }
                }
            }

            if (!TargetFound) {
                MsgBox("该项目非本地项目，在打开的窗口中未找到名为 '" . ProjectName . "' 的Citavi项目窗口。`n`n请手动在Citavi中打开此项目后再试。")
                ExitApp
            }
        }

        ; --- 后续逻辑 (如果窗口已找到并激活) ---
        if (TargetFound) {
            A_Clipboard := IdValue
            SwitchIME(0x04090409) ;切换英文，SwitchIME(00000804)是切换中文
            switch IdType {
                case "Annot":
                    Send("!M00R")
                    ShowTimedTrayTip("识别到AnnotID")
                case "Know":
                    Send("!M01R")
                    ShowTimedTrayTip("识别到KnowID")
                case "Ref":
                    Send("!M02R")
                    ShowTimedTrayTip("识别到RefID")
                default:
                    MsgBox("未知的Citavi ID类型: " . IdType)
                    ExitApp
            }
        }
    } else {
        errMsg := "未知的ahk协议或参数格式错误: " . UrlToProcess
        A_Clipboard := errMsg
        MsgBox(errMsg)
    }
} else {
    MsgBox("此脚本只能通过 ahk:// 协议调用。`n`n例如: ahk://citavi/goto?type=Ref&id=xxxxx&project=你的库名.ctv6&projectType=DesktopSQLite")
}

ExitApp
