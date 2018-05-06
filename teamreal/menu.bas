Attribute VB_Name = "menu"
Sub AddCommandbar()
    Dim i As Byte
    'For i = 0 To 6
    On Error Resume Next
    Application.CommandBars("用户菜单").Delete
    Application.CommandBars.Add "用户菜单", 1, , True

    Application.CommandBars("用户菜单").Visible = True
    With Application.CommandBars("用户菜单").Controls
         With .Add(1, , , , True)
            .Caption = "网站数据"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 10
            .OnAction = "网站数据更新"
        End With
        
        With .Add(1, , , , True)
            .Caption = "数据更新"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 11
            .OnAction = "数据更新"
        End With
        
        With .Add(1, , , , True)
            .Caption = "视图刷新"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 12
            .OnAction = "视图刷新"
        End With
        
    End With
End Sub


Sub DelCommandBars()
    On Error Resume Next
    Application.CommandBars("用户菜单").Delete
End Sub



