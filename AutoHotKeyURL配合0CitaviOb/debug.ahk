; debug.ahk (AutoHotkey v2.0 版本)

#SingleInstance Force ; v2.0 语法，没有逗号

; 准备一个变量来存储所有信息
AllInfo := "=== 调试信息 ===`n`n"
AllInfo .= "脚本完整路径: " . A_ScriptFullPath . "`n`n"
AllInfo .= "工作目录: " . A_WorkingDir . "`n`n"
AllInfo .= "收到的参数个数: " . A_Args.Length . "`n`n" ; v2.0 中用 A_Args.Length 获取参数个数

; 循环显示每一个参数
for index, value in A_Args ; v2.0 中用 for 循环遍历 A_Args 数组
{
    AllInfo .= "参数 " . index . ": " . value . "`n"
}

; 用一个文本框显示所有信息，方便复制
MsgBox(AllInfo) ; v2.0 中 MsgBox 是函数，需要用括号
Clipboard := AllInfo
ExitApp
