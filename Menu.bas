Attribute VB_Name = "Menu"
Sub AddCommandbars()
    Dim i As Byte
    'For i = 0 To 6
    On Error Resume Next
    Application.CommandBars("彩票分析").Delete
    Application.CommandBars.Add "彩票分析", 1, , True

    Application.CommandBars("彩票分析").Visible = True
    With Application.CommandBars("彩票分析").Controls
         With .Add(1, , , , True)
            .Caption = "网站数据"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 10
            .OnAction = "网站数据更新"
        End With
        
        With .Add(1, , , , True)
            .Caption = "初始"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 11
            .OnAction = "数据初始"
        End With
        
        With .Add(1, , , , True)
            .Caption = "更新"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 12
            .OnAction = "数据更新"
        End With
        With .Add(1, , , , True)
            .Caption = "模式计算"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 22
            .OnAction = "模式计算"
        End With
        With .Add(1, , , , True)
            .Caption = "历史数据"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 44
            .OnAction = "历史数据加载"
        End With
        
        With .Add(1, , , , True)
            .Caption = "查看当期数据"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 25
            .OnAction = "查看当期数据"
        End With
        
        With .Add(1, , , , True)
            .Caption = "查看全部赛事"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 26
            .OnAction = "查看全部信息"
        End With

        With .Add(1, , , , True)
            .Caption = "手工数据刷新"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 46
            .OnAction = "手工数据刷新"
        End With
        
        With .Add(1, , , , True)
            .Caption = "相同赔率比较"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 28
            .OnAction = "相同赔率比较"
        End With
        
        With .Add(1, , , , True)
            .Caption = "实力值"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 27
            .OnAction = "实力值计算"
        End With
        
        With .Add(1, , , , True)
            .Caption = "程序升级"
            .Visible = True
            .Style = msoButtonIconAndCaption
            .FaceId = 15
            .OnAction = "程序升级"
        End With
        
    End With
End Sub


Sub DelCommandBars()
    On Error Resume Next
    Application.CommandBars("彩票分析").Delete
End Sub



