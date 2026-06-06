Option Compare Database
Option Explicit

'确认
Public Sub Alert(ByVal MessageContent As String)
    MsgBox MessageContent, vbInformation + vbOKOnly, App.AppTitle
End Sub

'警告
Public Sub Warning(ByVal MessageContent As String)
    MsgBox MessageContent, vbCritical + vbOKOnly, App.AppTitle
End Sub

'确认+取消
Public Function Confirm(ByVal MessageContent As String) As Boolean
    If MsgBox(MessageContent, vbQuestion + vbOKCancel, App.AppTitle) = vbOK Then
        Confirm = True
    Else
        Confirm = False
    End If
End Function

'错误提示
Public Sub Error(ByRef ErrorObject As Object, Optional ByVal MessageContent As String)
    If Not IsMissing(MessageContent) Then
        MsgBox MessageContent, vbCritical + vbOKOnly, "系统错误"
    Else
        Select Case ErrorObject.Number
            Case 5
                MsgBox "过程调用无效", vbCritical + vbOKOnly, "系统错误"
            Case 6
                MsgBox "溢出", vbCritical + vbOKOnly, "系统错误"
            Case 7
                MsgBox "内存不足", vbCritical + vbOKOnly, "系统错误"
            Case 9
                MsgBox "数组下标越界", vbCritical + vbOKOnly, "系统错误"
            Case 13
                MsgBox "类型不匹配", vbCritical + vbOKOnly, "系统错误"
            Case 53
                MsgBox "文件未找到，请检查路径", vbCritical + vbOKOnly, "系统错误"
            Case 55
                MsgBox "文件已打开，请先关闭", vbCritical + vbOKOnly, "系统错误"
            Case 58
                MsgBox "文件已存在", vbCritical + vbOKOnly, "系统错误"
            Case 61
                MsgBox "磁盘空间不足，请清理空间", vbCritical + vbOKOnly, "系统错误"
            Case 76
                MsgBox "路径未找到", vbCritical + vbOKOnly, "系统错误"
            Case 70
                MsgBox "写入权限被拒绝", vbCritical + vbOKOnly, "系统错误"
            Case 91
                MsgBox "对象未创建或已被释放", vbCritical + vbOKOnly, "系统错误"
            Case 94
                MsgBox "字段值为 Null", vbCritical + vbOKOnly, "系统错误"
            Case 380
                MsgBox "属性值无效", vbCritical + vbOKOnly, "系统错误"
            Case 438
                MsgBox "对象不支持此属性或方法", vbCritical, "系统错误"
            Case 444
                MsgBox "方法调用无效", vbCritical + vbOKOnly, "系统错误"
            Case 445
                MsgBox "对象不支持此操作", vbCritical + vbOKOnly, "系统错误"
            Case 446
                MsgBox "未找到命名参数", vbCritical + vbOKOnly, "系统错误"
            Case 447
                MsgBox "参数未找到", vbCritical + vbOKOnly, "系统错误"
            Case 448
                MsgBox "未找到指定参数", vbCritical + vbOKOnly, "系统错误"
            Case 449
                MsgBox "参数值无效或属性不存在", vbCritical + vbOKOnly, "系统错误"
            Case 450
                MsgBox "参数数量或类型错误", vbCritical + vbOKOnly, "系统错误"
            Case 451
                MsgBox "对象不是集合", vbCritical + vbOKOnly, "系统错误"
            Case 452
                MsgBox "不支持的区域设置", vbCritical + vbOKOnly, "系统错误"
            Case 453
                MsgBox "找不到 DLL 入口点", vbCritical + vbOKOnly, "系统错误"
            Case 458
                MsgBox "变量类型不支持", vbCritical + vbOKOnly, "系统错误"
            Case 3001
                MsgBox "参数类型错误", vbCritical + vbOKOnly, "系统错误"
            Case 3006
                MsgBox "数据库已被独占打开，请关闭其他程序", vbCritical + vbOKOnly, "系统错误"
            Case 3011
                MsgBox "表或查询不存在", vbCritical + vbOKOnly, "系统错误"
            Case 3022
                MsgBox "主键重复", vbCritical + vbOKOnly, "系统错误"
            Case 3031
                MsgBox "密码错误", vbCritical + vbOKOnly, "系统错误"
            Case 3033
                MsgBox "没有操作权限", vbCritical + vbOKOnly, "系统错误"
            Case 3048
                MsgBox "无法锁定数据库", vbCritical + vbOKOnly, "系统错误"
            Case 3049
                MsgBox "数据库可能已损坏，请压缩修复", vbCritical + vbOKOnly, "系统错误"
            Case 3050
                MsgBox "文件已被锁定", vbCritical + vbOKOnly, "系统错误"
            Case 3051
                MsgBox "网络连接已断开", vbCritical + vbOKOnly, "系统错误"
            Case 3061
                MsgBox "参数个数错误", vbCritical + vbOKOnly, "系统错误"
            Case 3065
                MsgBox "无法执行查询", vbCritical + vbOKOnly, "系统错误"
            Case 3070
                MsgBox "记录集不支持更新", vbCritical + vbOKOnly, "系统错误"
            Case 3078
                MsgBox "表或查询不存在", vbCritical + vbOKOnly, "系统错误"
            Case 3075
                MsgBox "SQL 语法错误", vbCritical + vbOKOnly, "系统错误"
            Case 3085
                MsgBox "字段未定义", vbCritical + vbOKOnly, "系统错误"
            Case 3086
                MsgBox "无法删除记录", vbCritical + vbOKOnly, "系统错误"
            Case 3091
                MsgBox "外键约束冲突", vbCritical + vbOKOnly, "系统错误"
            Case 3112
                MsgBox "无法读取记录", vbCritical + vbOKOnly, "系统错误"
            Case 3113
                MsgBox "无法写入记录", vbCritical + vbOKOnly, "系统错误"
            Case 3144
                MsgBox "UPDATE 语句语法错误", vbCritical + vbOKOnly, "系统错误"
            Case 3155
                MsgBox "INSERT INTO 语法错误", vbCritical + vbOKOnly, "系统错误"
            Case 3162
                MsgBox "必填字段不能为空", vbCritical + vbOKOnly, "系统错误"
            Case 3163
                MsgBox "字段值类型不匹配", vbCritical + vbOKOnly, "系统错误"
            Case 3167
                MsgBox "记录已被删除", vbCritical + vbOKOnly, "系统错误"
            Case 3188
                MsgBox "记录被其他用户锁定", vbCritical + vbOKOnly, "系统错误"
            Case 3197
                MsgBox "记录被其他用户修改", vbCritical + vbOKOnly, "系统错误"
            Case 3200
                MsgBox "无法追加记录", vbCritical + vbOKOnly, "系统错误"
            Case 3201
                MsgBox "无法删除或修改该记录", vbCritical + vbOKOnly, "系统错误"
            Case 3219
                MsgBox "当前状态下无效操作", vbCritical + vbOKOnly, "系统错误"
            Case 3251
                MsgBox "当前记录集不支持此操作", vbCritical + vbOKOnly, "系统错误"
            Case 3256
                MsgBox "数据已被其他用户更改", vbCritical + vbOKOnly, "系统错误"
            Case 3265
                MsgBox "字段不存在", vbCritical + vbOKOnly, "系统错误"
            Case 3314
                MsgBox "必填字段不能为 Null", vbCritical + vbOKOnly, "系统错误"
            Case 3376
                MsgBox "记录正被其他用户锁定", vbCritical + vbOKOnly, "系统错误"
            Case 3420
                MsgBox "数据类型转换失败", vbCritical + vbOKOnly, "系统错误"
            Case Else
                MsgBox "错误 [" & ErrorObject.Number & "] : " & ErrorObject.Description, vbCritical + vbOKOnly, "系统错误"
        End Select
    End If
    ErrorObject.Clear
End Sub
