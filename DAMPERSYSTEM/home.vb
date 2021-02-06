Imports SldWorks
Imports Microsoft.Office.Interop.Excel
Imports MySql.Data.MySqlClient
Public Class home
    Dim pi As Double = 3.141592653
    Dim mysql As MySqlConnection
    Dim swapp As SldWorks.SldWorks
    Dim part As ModelDoc2
    Dim asm As AssemblyDoc
    Dim kong As Feature
    Public xlapp As Application ''引用Microsoft excel和Microsoft office类型库
    Public xlBook As Workbook
    Public xlSheet As Worksheet ''然后创建对象
    Public wt(3) As Double '外筒参数数组
    Public hsg As Double '活塞杆数组
    Public hs As Double '活塞数组
    Dim alpha As Double '阻尼系数
    Dim I As Integer
    Dim aProcesses() As Process = Process.GetProcesses
    Dim XLAPPpid As Integer
    Dim parttitle As String '文件名
    Dim feature As Feature ''拉伸特征
    Dim scale As Double ''缩放比例
    Dim Parachaval As Double ''过程参数
    'Dim mysqlconnect As MySqlConnection ''定义mysql连接
    'Dim mycommand As MySqlCommand ''定义mysql命令
    'Dim reader As MySqlDataReader ''定义数据流
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        swapp = CreateObject("Sldworks.Application")
        part = swapp.ActiveDoc
        parttitle = part.GetTitle
        swapp.CloseDoc(parttitle)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        partsub.Text = "外筒生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\装配体外筒.png")
        Me.Hide()
        partsub.parttype = 1
        partsub.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        partsub.Text = "活塞杆生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\装配体活塞杆.png")
        Me.Hide()
        partsub.parttype = 2
        partsub.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        partsub.Text = "活塞生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\装配体活塞.png")
        Me.Hide()
        partsub.parttype = 3
        partsub.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        partpanel.Visible = True
        assemblepanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        partpanel.Visible = False
        assemblepanel.Visible = True
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = True
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        setdamper.Visible = False
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = True
        drawingpanel.Visible = False
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        For Each myprocess As Process In Process.GetProcesses
            If InStr(myprocess.ProcessName, "SLDWORKS") Then
                myprocess.Kill()
            End If
        Next
        Dim p As Process() = Process.GetProcessesByName("EXCEL")
        For Each pr As Process In p
            pr.Kill()
        Next
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        partpanel.Visible = False
        assemblepanel.Visible = False
        morepanel.Visible = True
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        partpanel.Visible = True
        assemblepanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim matlab As Object
        Dim result As String
        matlab = CreateObject("matlab.application")
        result = matlab.execute("addpath(genpath('D:\阻尼器布置算法'))")
        result = matlab.execute("ModelInformation;")
        MsgBox(result)
        result = matlab.execute("GenerateAB0;")
        result = matlab.execute("HPDSolve;")
        result = matlab.execute("d0max = max(abs(d'));a0max=max(abs(a'));")
        result = matlab.execute("UnfixedQuantityDamperOptimization;")
        MsgBox(result)
        Try
            Dim p As Process() = Process.GetProcessesByName("MATLAB")
            For Each pr As Process In p
                pr.Kill()
            Next
        Catch ex As Exception
            MessageBox.Show("计算完成")
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text IsNot "" And TextBox2.Text IsNot "" And
            TextBox3.Text IsNot "" And TextBox4.Text IsNot "" Then
            Excel()
        Else
            MsgBox("未输入值")
            Exit Sub
        End If
        'wt(0) = 130 / 1000 '外径  
        'wt(1) = 100 / 1000 '内径
        'wt(2) = 160 / 1000 '长度
        'hs = 78.12 / 1000 '活塞直径
        'hsg = 30 / 1000 ''活塞杆直径(接拉头螺纹直径20-25mm)
        wt(0) = wt(0) / 1000
        wt(1) = wt(1) / 1000
        wt(2) = 400 / 1000
        hs = hs / 1000
        hsg = hsg / 1000
        Createwaitong()
        'createcylinderhead()
        'createspacerpiece()
        'Createrod()
        'createrodbearing()
        'createpiston()
        'createthreadedflange()
        'createexrod()
        'createextube()
        'Createcapeex()
        'Createsocconrodend()
        'Createterend()
        'Createreaear()
        'Createothers()
    End Sub
    Public Function Excel()
        xlapp = CreateObject("Excel.Application")  ''创建EXCEL对象
        xlBook = xlapp.Workbooks.Open("D:\POST-GRA\研究生大论文\零件库\筒式-间隙-层间用-设计参数阵.xlsx")
        ''打开已经存在的EXCEL工件簿文件
        alpha = CType(TextBox2.Text, Double)
        Select Case alpha''阻尼系数类型
            Case 0.2
                xlSheet = xlBook.Worksheets("α=0.20")
            Case 0.25
                xlSheet = xlBook.Worksheets("α=0.25")
            Case 0.3
                xlSheet = xlBook.Worksheets("α=0.30")
            Case Else
                RichTextBox1.Text = "未找到参数"
                Closeexcel()
                Exit Function
        End Select
        ReDim wt(2)
        Dim textbox1text As Double = CType(TextBox1.Text, Double）
        Select Case True''阻尼系数类型
            Case textbox1text <= 150
                wt(0) = 110 '直径
                wt(1) = 80
                wt（2） = 15
                hsg = 30
            Case textbox1text > 150 And textbox1text <= 250
                wt(0) = 130 '直径
                wt(1) = 100
                wt（2） = 15
                hsg = 40
            Case textbox1text > 250 And textbox1text <= 350
                wt(0) = 150 '直径
                wt(1) = 110
                wt（2） = 20
                hsg = 45
            Case textbox1text > 360 And textbox1text <= 450
                wt(0) = 166 '直径
                wt(1) = 126
                wt（2） = 20
                hsg = 50
            Case textbox1text > 450 And textbox1text <= 650
                wt(0) = 200 '直径
                wt(1) = 150
                wt（2） = 25
                hsg = 60
            Case Else
                RichTextBox1.Text = "未找到参数"
                Closeexcel()
                Exit Function
        End Select
        I = 0
        For Each sh In xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(114, 49))
            If sh.row = 114 Then
                RichTextBox1.Text = "未找到参数"
                Closeexcel()
                Exit For
            Else
                If IsNumeric(sh.value) Then
                    If Math.Abs(sh.value - CType（TextBox1.Text, Double）) < 0.1 And
                Math.Abs(xlSheet.Cells(sh.row, sh.column + 1).Value - CType（TextBox3.Text, Double）) < 0.1 Then
                        hs = xlSheet.Cells(sh.row, sh.column + 4).Value
                        RichTextBox1.Text = "筒外径：" & wt（0） & vbCrLf &
                            "筒内径：" & wt（1） & vbCrLf &
                            "活塞直径：" & hs & vbCrLf &
                            "活塞杆直径：" & hsg
                        Closeexcel()
                        GC.Collect()
                        Exit For
                    End If
                End If
            End If
        Next
        GC.Collect()
    End Function
    Public Function Closeexcel()
        If xlBook IsNot Nothing Then
            xlBook.Close()
            xlapp.Quit()
            xlapp = Nothing
        End If
    End Function
    Public Function Createwaitong() ''外筒（完成）
        ''创建进程可视化
        swapp = CreateObject("Sldworks.Application")
        'swapp.CloseAllDocuments(True)

        ''创建新零件
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        swapp.Visible = True
        part = swapp.ActiveDoc
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.SketchManager.InsertSketch(True)
        part.ClearSelection2(True)
        part.SketchManager.CreateCircle(0#, 0#, 0#, wt(0) / 2, 0, 0#)
        part.SketchManager.CreateCircle(0#, 0#, 0#, wt(1) / 2, 0, 0#)
        part.FeatureManager.FeatureExtrusion3(True, True, False, 0, 0,
                                                wt（2） / 2, 0, False, False, False,
                                                False, 0, 0, True, True,
                                                False, False, True, False, False,
                                                0, 0, 0)
        part.ClearSelection2(True)
        part.Extension.SelectByID2("", "FACE", (wt(0) + wt(1)) / 4, 0, wt(2) / 2, False, 0, Nothing, 0)
        part.SketchManager.InsertSketch(True)
        part.SketchManager.CreateCircle(0, 0, wt（2) / 2, (wt(0) + wt(1)) / 4, 0, wt（2) / 2)
        part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.01, 0.01, False,
                                        False, False, False, 0, 0,
                                        False, False, False, False, False, True, True, True, True, False,
                                         0, 0, False, False)
        part.ShowNamedView2("*前视", 1)
        part.Extension.SelectByID2("", "", (wt(0) + wt(1)) / 4, 0, wt(2) / 2, False, 0, Nothing, 0)
        part.FeatureManager.InsertCosmeticThread2(1, (wt(0) + wt(1)) / 2, 0, "140")
        part.Extension.SelectByID2("", "", wt(1) / 2, 0, wt(2) / 2 - 0.01, False, 0, Nothing, 0)
        part.FeatureManager.InsertFeatureChamfer(7, 1, 0.013, 7.5 * pi / 180, 0, 0, 0, 0)
        part.ShowNamedView2("", 7)
        part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.FeatureManager.InsertRefPlane(8, wt(0) / 2, 0, 0, 0, 0)
        Dim swWzdHole As WizardHoleFeatureData2
        swWzdHole = part.FeatureManager.CreateDefinition(SwConst.swFeatureNameID_e.swFmHoleWzd)
        part.Extension.SelectByID2("", "FACE", wt(0) / 2, 0, wt(2) / 4, False, 0, Nothing, 0)
        part.FeatureManager.HoleWizard5(4, 1, 42, "M10x1.0", 2, 0.009, 0.02, 0.02, 0, 0, 0, 0, 0, 0, 2, 0,
                                        0, -1, -1, -1, "", False, True, True, True, True, False)
        part.Extension.SelectByID2("凸台-拉伸1", "BODYFEATURE", 0, 0, 0, False, 1, Nothing, 0) ''1是镜像特征
        part.Extension.SelectByID2("M10x1.0 螺纹孔1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        part.Extension.SelectByID2("切除-拉伸1"， "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 2, Nothing, 0) ''2是镜像基准
        part.FeatureManager.InsertMirrorFeature(False, False, False, False)
        part.ShowNamedView2("", 2)
        part.Extension.SelectByID2("", "", wt(1) / 2, 0, -wt(2) / 2 + 0.01, False, 0, Nothing, 0)
        part.FeatureManager.InsertFeatureChamfer(7, 1, 0.013, 7.5 * pi / 180, 0, 0, 0, 0)
        part.Extension.SelectByID2("M10x1.0 螺纹孔1", "BODYFEATURE", 0, 0, 0, False, 1, Nothing, 0)
        ''修正螺纹线参数，设置螺纹线为1m，自动生成
        feature = part.SelectionManager.GetSelectedObject5(1)
        swWzdHole = feature.GetDefinition
        swWzdHole.ThreadDepth = 1
        feature.ModifyDefinition(swWzdHole, part, Nothing)
        part.ShowNamedView2("", 7)
        part.EditRebuild3()
    End Function
    Public Function Createcylinderhead() ''端盖（完成）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Cylinider head.SLDPRT",
                              1, 0, "", 0, 0)
        ''主动尺寸
        Parachaval = part.Parameter("D7@Sketch1").SYSTEMVALUE - hsg / 2
        part.Parameter("D7@Sketch1").SYSTEMVALUE = hsg / 2 ''活塞杆直径
        part.Parameter("D28@Sketch1").SYSTEMVALUE = wt(1) / 2 ''内筒内径
        part.Parameter("D29@Sketch1").SYSTEMVALUE = wt(1) / 2 ''内筒内径
        ''随动尺寸
        part.Parameter("D6@Sketch1").SYSTEMVALUE = part.Parameter("D6@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D10@Sketch1").SYSTEMVALUE = part.Parameter("D10@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D9@Sketch1").SYSTEMVALUE = part.Parameter("D9@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D25@Sketch1").SYSTEMVALUE = part.Parameter("D25@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D13@Sketch1").SYSTEMVALUE = part.Parameter("D13@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D11@Sketch1").SYSTEMVALUE = part.Parameter("D11@Sketch1").SYSTEMVALUE - Parachaval
        ''孔尺寸随动
        part.Parameter("D2@Sketch6").SYSTEMVALUE = part.Parameter("D2@Sketch6").SYSTEMVALUE - Parachaval
        part.EditRebuild3()
    End Function
    Public Function Createspacerpiece() ''垫片（完成）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Spacer piece.SLDPRT", 1, 0, "", 0, 0)
        part.Parameter("D2@Sketch2").systemvalue = wt(1) / 2
        part.Parameter("D5@Sketch2").systemvalue = wt(1) / 2
        part.Parameter("D18@Sketch2").systemvalue = wt(1) / 2 ''旋转内径
        part.Parameter("D11@Sketch2").systemvalue = part.Parameter("D11@Sketch2").systemvalue - (80 / 1000 - wt(1) / 2)
        part.Parameter("D14@Sketch2").systemvalue = part.Parameter("D14@Sketch2").systemvalue - (80 / 1000 - wt(1) / 2)
        part.Parameter("D3@Sketch2").systemvalue = part.Parameter("D3@Sketch2").systemvalue - (80 / 1000 - wt(1) / 2)
        part.Parameter("D10@Sketch2").systemvalue = part.Parameter("D10@Sketch2").systemvalue - (80 / 1000 - wt(1) / 2)
        ''修改特征的方法：1、选中特征；2、利用selectionmgr获取对象；、3、set对象属性；4、修改modify定义
        part.Extension.SelectByID2("Cut-Extrude2", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        feature = part.SelectionManager.GetSelectedObject5(1) ''获取选中实体的特征对象
        Dim extrudedata As ExtrudeFeatureData2
        extrudedata = feature.GetDefinition ''获取特征的定义
        extrudedata.SetEndCondition(1, 2) ''修改特征定义
        feature.ModifyDefinition(extrudedata, part, Nothing) ''修改到特征中
        part.EditRebuild3()
    End Function
    Public Function Createrod() ''活塞杆（连接处形状不变）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Rod.SLDPRT",
                              1, 0, "", 0, 0)
        Parachaval = part.Parameter("D2@Sketch1").systemvalue - hsg / 2
        part.Parameter("D2@Sketch1").systemvalue = hsg / 2
        part.Parameter("D10@Sketch1").systemvalue = part.Parameter("D10@Sketch1").systemvalue - Parachaval + 0.01
        part.Parameter("D9@Sketch1").systemvalue = part.Parameter("D9@Sketch1").systemvalue - Parachaval
        part.Parameter("D4@Sketch1").systemvalue = part.Parameter("D4@Sketch1").systemvalue - Parachaval
        part.EditRebuild3()
    End Function
    Public Function Createrodbearing() ''连杆轴承(内圈D-35是活塞杆直径，螺纹是端盖内螺纹)
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Rod bearing.SLDPRT",
                              1, 0, "", 0, 0)
        Parachaval = part.Parameter("D24@Sketch1").systemvalue - hsg / 2
        part.Parameter("D24@Sketch1").systemvalue = hsg / 2
        part.Parameter("D27@Sketch1").systemvalue = hsg / 2
        ''等比例外延
        part.Parameter("D18@Sketch1").systemvalue = part.Parameter("D18@Sketch1").systemvalue - Parachaval
        part.Parameter("D16@Sketch1").systemvalue = part.Parameter("D16@Sketch1").systemvalue - Parachaval
        part.Parameter("D7@Sketch1").systemvalue = part.Parameter("D7@Sketch1").systemvalue - Parachaval
        part.Parameter("D13@Sketch1").systemvalue = part.Parameter("D13@Sketch1").systemvalue - Parachaval
        part.Parameter("D12@Sketch1").systemvalue = part.Parameter("D12@Sketch1").systemvalue - Parachaval
        part.Parameter("D2@Sketch1").systemvalue = part.Parameter("D2@Sketch1").systemvalue - Parachaval
        part.Parameter("D4@Sketch1").systemvalue = part.Parameter("D4@Sketch1").systemvalue - Parachaval
        part.Parameter("D3@Sketch1").systemvalue = part.Parameter("D3@Sketch1").systemvalue - Parachaval
        part.EditRebuild3()
    End Function
    Public Function Createpiston() ''活塞
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\piston.SLDPRT",
                              1, 0, "", 0, 0)
        part.Parameter("D1@Sketch1").systemvalue = hs ''活塞直径（外径）
        Parachaval = part.Parameter("D16@Sketch3").systemvalue - hsg / 2
        ''内径
        part.Parameter("D16@Sketch3").systemvalue = hsg / 2
        part.Parameter("D11@Sketch3").systemvalue = part.Parameter("D11@Sketch3").systemvalue - Parachaval
        part.Parameter("D10@Sketch3").systemvalue = part.Parameter("D10@Sketch3").systemvalue - Parachaval
        part.EditRebuild3()
    End Function
    Public Function Createthreadedflange() ''螺纹法兰（完成）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Threaded flange.SLDPRT",
                              1, 0, "", 0, 0)
        Parachaval = part.Parameter("D1@Sketch1").systemvalue - (70 / 1000 + hsg) ''cylinderhead.(d6-(d7-hsg/2))
        part.Parameter("D1@Sketch1").systemvalue = part.Parameter("D1@Sketch1").systemvalue - Parachaval
        part.Parameter("D2@Sketch1").systemvalue = part.Parameter("D1@Sketch1").systemvalue + 30 / 1000
        part.Parameter("D2@Sketch3").systemvalue = part.Parameter("D2@Sketch3").systemvalue - Parachaval / 2
        part.Extension.SelectByID2("Cosmetic Thread1", "CTHREAD", 0, 0, 0, False, 0, Nothing, 0)
        ''指定数据类型
        Dim thread As CosmeticThreadFeatureData
        ''获取特征
        feature = part.SelectionManager.GetSelectedObject5(1)
        ''获取特征定义
        thread = feature.GetDefinition
        ''修改特征定义
        thread.Diameter = thread.Diameter - Parachaval
        ''修改特征
        feature.ModifyDefinition(thread, part, Nothing)
        part.EditRebuild3()
    End Function
    Public Function Createexrod()
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Extension rod.SLDPRT",
                              1, 0, "", 0, 0)
        Parachaval = part.Parameter("D1@Sketch2").systemvalue - hsg / 2
        part.Parameter("D1@Sketch2").systemvalue = hsg / 2
        part.Parameter("D6@Sketch2").systemvalue = 5 / 1000
        part.Parameter("D9@Sketch2").systemvalue = 6.25 / 1000
        part.EditRebuild3()
    End Function
    Public Function Createextube() ''伸长筒（完成
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Extension tube.SLDPRT",
                              1, 0, "", 0, 0)
        ''EXTENSION TUBE.D1=CYLINDERHEAD.D6=70-(35-HSG/2)
        Parachaval = part.Parameter("D1@Sketch1").SYSTEMVALUE - (70 / 1000 - (35 / 1000 - hsg / 2))
        part.Parameter("D1@Sketch1").SYSTEMVALUE = (70 / 1000 - (35 / 1000 - hsg / 2))
        part.Parameter("D2@Sketch1").SYSTEMVALUE = part.Parameter("D2@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D9@Sketch1").SYSTEMVALUE = part.Parameter("D9@Sketch1").SYSTEMVALUE - Parachaval
        part.Extension.SelectByID2("Cosmetic Thread1", "CTHREAD", 0, 0, 0, False, 0, Nothing, 0)
        ''指定数据类型
        Dim thread As CosmeticThreadFeatureData
        ''获取特征
        feature = part.SelectionManager.GetSelectedObject5(1)
        ''获取特征定义
        thread = feature.GetDefinition
        ''修改特征定义
        thread.Diameter = thread.Diameter - Parachaval * 2
        ''修改特征
        feature.ModifyDefinition(thread, part, Nothing)
        part.EditRebuild3()
    End Function
    Public Function Createcapeex() ''角延伸管(完成)
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Cape extension tube.SLDPRT",
                              1, 0, "", 0, 0)
        ''createextube.d9-20.6=33.4+hsg/2
        part.Parameter("D1@Sketch1").SYSTEMVALUE = 66.8 / 1000 + hsg
        part.EditRebuild3()
    End Function
    Public Function Createsocconrodend() ''套筒连接杆端(完成)
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Socket connection rod end.SLDPRT",
                              1, 0, "", 0, 0)
        part.Parameter("D3@Sketch3").SYSTEMVALUE = 10 / 1000
        part.Parameter("D4@Sketch3").SYSTEMVALUE = 11.5 / 1000
        part.Parameter("D1@Sketch1").SYSTEMVALUE = part.Parameter("D1@Sketch1").SYSTEMVALUE - 18.75 / 1000 * 2
        part.EditRebuild3()
    End Function
    Public Function Createterend() ''终端（完成）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Terminal end.SLDPRT",
                              1, 0, "", 0, 0)
        part.Parameter("D2@Sketch2").SYSTEMVALUE = 62.5 / 1000
        part.Extension.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        feature = part.SelectionManager.GETSELECTEDOBJECT5(1)
        Dim EXTRU As ExtrudeFeatureData2
        EXTRU = feature.GetDefinition
        EXTRU.SetDepth(1, 50 / 1000)
        feature.ModifyDefinition(EXTRU, part, Nothing)
        part.EditRebuild3()
    End Function
    Public Function Createreaear() ''后耳座（完成）
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Rear ear.SLDPRT",
                              1, 0, "", 0, 0)
        part.Parameter("D1@Sketch1").SYSTEMVALUE = 100 / 1000
        part.Parameter("D3@Sketch1").SYSTEMVALUE = 50 / 1000
        part.EditRebuild3()
    End Function
    Public Function Createothers() ''密封件采用缩放形式
        swapp = CreateObject("Sldworks.Application")
        swapp.Visible = True
        '''''''''''''''''''''''''''''''''
        ''''GP6901600-T47 (0).sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\GP6901600-T47 (0).sldprt",
                              1, 0, "", 0, 0)
        ''D1=SPACESPIECE.D11
        Parachaval = part.Parameter("D1@Sketch2").SYSTEMVALUE - (wt(1) / 2) + 2.5 / 1000
        part.Parameter("D1@Sketch2").SYSTEMVALUE = wt(1) / 2 - 2.5 / 1000
        part.Parameter("D2@Sketch2").SYSTEMVALUE = part.Parameter("D1@Sketch2").SYSTEMVALUE + 0.3 / 1000
        part.Parameter("D1@Sketch1").SYSTEMVALUE = part.Parameter("D1@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D2@Sketch1").SYSTEMVALUE = part.Parameter("D2@Sketch1").SYSTEMVALUE - Parachaval
        part.EditRebuild3()
        ''''''''''''''''''''''''''''''''''''''''''''
        ''''ORID14500-N5 - (00).sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\ORID14500-N5 - (00).sldprt",
                              1, 0, "", 0, 0)

        ''D1=CYLINDERHEAD.D28(实际)-4.3+2.15=WT(1)/2-4.3+2.15
        part.Parameter("D1@Sketch1").SYSTEMVALUE = wt(1) / 2 - 4.3 / 1000 + 2.15 / 1000
        part.Parameter("D2@Sketch1").SYSTEMVALUE = wt(1) / 2 - 4.3 / 1000 + 2.15 / 1000
        part.EditRebuild3()
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''GR6900700-T47 (0).sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\GR6900700-T47 (0).sldprt",
                              1, 0, "", 0, 0)
        Parachaval = part.Parameter("D2@Sketch2").SYSTEMVALUE - hsg / 2
        part.Parameter("D2@Sketch2").SYSTEMVALUE = hsg / 2
        part.Parameter("D1@Sketch2").SYSTEMVALUE = part.Parameter("D1@Sketch2").SYSTEMVALUE - Parachaval
        part.Parameter("D2@Sketch1").SYSTEMVALUE = part.Parameter("D2@Sketch1").SYSTEMVALUE - Parachaval
        part.Parameter("D1@Sketch1").SYSTEMVALUE = part.Parameter("D1@Sketch1").SYSTEMVALUE - Parachaval
        part.EditRebuild3()
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''head back-up ring.sldprt(头部支承环)
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\head back-up ring.SLDPRT",
                              1, 0, "", 0, 0)
        '''' CYLINDERHEAD.D28-4.3
        part.Parameter("D1@Sketch1").SYSTEMVALUE = wt(1) / 2 - 4.3 / 1000
        part.Parameter("D2@Sketch1").SYSTEMVALUE = part.Parameter("D1@Sketch1").SYSTEMVALUE + 8.6 / 1000
        part.EditRebuild3()
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''ORAR00151.sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\ORAR00151.sldprt",
1, 0, "", 0, 0)
        scale = (40.5 - (35 - hsg / 2)) / 40.5
        part.FeatureManager.InsertScale(0, True, scale, scale, scale)
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''WAP100700-N9T60 (0).sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\WAP100700-N9T60 (0).sldprt",
1, 0, "", 0, 0)
        scale = hsg / 2 / 35 * 1000
        part.FeatureManager.InsertScale(0, True, scale, scale, scale)
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''RSK3007001 (01).sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\RSK3007001 (01).sldprt",
1, 0, "", 0, 0)
        scale = hsg / 2 / 35 * 1000
        part.FeatureManager.InsertScale(0, True, scale, scale, scale)
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''RSK3007002.sldprt
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\RSK3007002.sldprt",
1, 0, "", 0, 0)
        scale = hsg / 2 / 35 * 1000
        part.FeatureManager.InsertScale(0, True, scale, scale, scale)
        '''''''''''''''''''''''''''''''''''''''''''''''
        ''''RL16N0700-Z20 (010)
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\RL16N0700-Z20 (010)",
1, 0, "", 0, 0)
        scale = hsg / 2 / 35 * 1000
        part.FeatureManager.InsertScale(0, True, scale, scale, scale)
    End Function
End Class
