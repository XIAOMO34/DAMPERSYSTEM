Imports SldWorks
Imports Microsoft.Office.Interop.Excel
Imports MySql.Data.MySqlClient
Public Class home ''kenan
    Dim pi As Double = 3.141592653
    Dim mysql As MySqlConnection
    Dim swapp As SldWorks.SldWorks
    Dim part As ModelDoc2
    Dim asm As AssemblyDoc
    Dim kong As Feature
    Public xlapp As Application ''引用Microsoft excel和Microsoft office类型库
    Public xlBook As Workbook
    Public xlSheet As Worksheet ''然后创建对象
    Dim wt(3) As Double '外筒参数数组
    Dim hsg As Double '活塞杆数组
    Dim hs As Double '活塞数组
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
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs)
        partpanel.Visible = True
        assemblepanel.Visible = False
        morepanel.Visible = False
    End Sub
    Private Sub BunifuFlatButton2_Click(sender As Object, e As EventArgs)
        partpanel.Visible = False
        assemblepanel.Visible = True
        morepanel.Visible = False
    End Sub
    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs)
        partpanel.Visible = False
        assemblepanel.Visible = False
        morepanel.Visible = True
    End Sub
    Private Sub BunifuImageButton1_Click_1(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub BunifuFlatButton1_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        partpanel.Visible = True
        assemblepanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub BunifuImageButton1_Click_2(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub BunifuFlatButton2_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton2.Click
        partpanel.Visible = False
        assemblepanel.Visible = True
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub BunifuFlatButton3_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        partpanel.Visible = False
        assemblepanel.Visible = False
        morepanel.Visible = True
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub Buttonduantou_Click_1(sender As Object, e As EventArgs) Handles Buttonduantou.Click
        Me.Hide()
        partsub.parttype = 1
        partsub.Show()
    End Sub
    Private Sub BunifuFlatButton10_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton10.Click
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = True
    End Sub
    Private Sub BunifuFlatButton13_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton13.Click
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

    Private Sub BunifuFlatButton11_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton11.Click
        setdamper.Visible = False
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub BunifuFlatButton12_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton12.Click
        assemblepanel.Visible = False
        partpanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = True
        drawingpanel.Visible = False
    End Sub
    ''工程图生成
    Private Sub BunifuFlatButton17_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton17.Click
        Dim ZhuView As Object
        Dim ZhuoView As Object
        Dim FuView As Object
        Dim myDisplayDim As Object
        swapp = CreateObject("sldworks.application")
        swapp.CloseAllDocuments(True)
        swapp.Visible = True
        part = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\gb_a4.drwdot", 12, 0, 0)
        ''通过模板创建新工程图
        ZhuView = part.CreateDrawViewFromModelView3("C:\Users\LJX\Desktop\零件1.SLDPRT", "*前视", 0.08, 0.16, 0)
        ''通过指定零件创建某视图的工程图
        part.Extension.SelectByID2("", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '工程图的类型为DRAWINGVIEW
        ZhuoView = part.CreateUnfoldedViewat3(0.2, 0.16, 0, False)
        ''通过已有视图创建新视图,false表示对齐，true不对齐
        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        FuView = part.CreateUnfoldedViewAt3(0.08, 0.1, 0, False)
        part.ClearSelection2(True)
        part.Extension.SelectByRay(0.0835202743534853, 0.154815383481483, 250, 0, 0, -1, 0.00105527915159568, 1, False, 0, 0)
        myDisplayDim = part.AddDimension2(0.12, 0.17, 0)
    End Sub

    Private Sub Buttonhuosaigan_Click(sender As Object, e As EventArgs) Handles Buttonhuosaigan.Click
        partsub.Text = "小头端盖生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\xtdg.JPG")
        Me.Hide()
        partsub.parttype = 4
        partsub.Show()
    End Sub

    Private Sub Buttonneitong_Click(sender As Object, e As EventArgs) Handles Buttonneitong.Click
        partsub.Text = "大头端盖生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\dtdg1.JPG")
        Me.Hide()
        partsub.parttype = 3
        partsub.Show()
    End Sub

    Private Sub Buttonzunitong_Click(sender As Object, e As EventArgs) Handles Buttonzunitong.Click
        partsub.Text = "外筒生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\wt方形.png")
        Me.Hide()
        partsub.parttype = 2
        partsub.Show()
    End Sub

    Private Sub BunifuFlatButton5_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton5.Click
        'Me.Hide()
        Dim featruemgr As FeatureManager
        Dim datouduangai As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\大头端盖.SLDPRT"
        Dim xiaotouduangai As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\小头端盖.SLDPRT"
        Dim latou As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\拉头.SLDPRT"
        Dim jietou As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\接头.SLDPRT"
        Dim waitong As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\外筒.SLDPRT"
        Dim huosaigan As String =
            "C:\Users\LJX\Desktop\装配练习2019-10-14\活塞杆.SLDPRT"
        Dim title As String
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        swapp.Visible = True
        part = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\gb_assembly.asmdot", 0,
                                 0, 0)
        title = part.GetTitle
        asm = part
        Addcomponent(waitong)
        Addcomponent(huosaigan)
        Addcomponent(datouduangai)
        Addcomponent(xiaotouduangai)
        swapp.OpenDoc6(latou, 1, 32, "", 2, 2)
        asm.AddComponent5(latou, 0, "", False, "", 1, 0, 0)
        swapp.CloseDoc(latou)
        Addcomponent(jietou)
        part.Extension.SelectByID2("右视基准面@大头端盖-1@" & title, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("右视基准面@外筒-1@" & title, "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        asm.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0,
                     0, 0, False, False, 0, 0)
        part.ShowNamedView2("*左视", 3)
        part.Extension.SelectByID2("右视基准面@接头-1@" & title, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("", "FACE", -0.153, 0.054, 0, True, 0, Nothing, 0)
        asm.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0,
                     0, 0, False, False, 0, 0)
        part.Extension.SelectByID2("", "FACE", 0.153, 0.05, 0, False, 0, Nothing, 0)
        part.ShowNamedView2("*左视", 3)
        part.Extension.SelectByID2("", "FACE", 1 - 0.28, 0.035, 0, True, 0, Nothing, 0)
        part.ShowNamedView2("*", 7)
        asm.AddMate5(0, 1, False, 0.001, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, False, False, 0, 1)
    End Sub
    Function Addcomponent(a As String)
        swapp.OpenDoc6(a, 1, 32, "", 2, 2)
        asm.AddComponent5(a, 0, "", False, "", 0, 0, 0)
        swapp.CloseDoc(a)
        Return 0
    End Function

    Private Sub BunifuFlatButton4_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        partsub.Text = "活塞杆生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\hsg.JPG")
        Me.Hide()
        partsub.parttype = 5
        partsub.Show()
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

    End Sub

    Private Sub BunifuFlatButton7_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton7.Click
        partsub.Text = "拉头生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\lt.JPG")
        Me.Hide()
        partsub.parttype = 6
        partsub.Show()
    End Sub

    Private Sub BunifuFlatButton14_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub BunifuFlatButton15_Click(sender As Object, e As EventArgs)
        ''创建进程可视化
        swapp = CreateObject("Sldworks.Application")
        ''swapp.CloseAllDocuments(True)
        ''创建新零件
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        swapp.Visible = True
        part = swapp.ActiveDoc
        part.Extension.SelectByID2("前视基准面"， "PLANE"， 0， 0, 0， False， 0， Nothing， 0)
        part.SketchManager.InsertSketch(True)
        part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.07)
        part.FeatureManager.FeatureExtrusion3(True, True, False, 0, 0,
                                                0.02, 0, False, False, False,
                                                False, 0, 0, True, True,
                                                False, False, True, False, False,
                                                0, 0, 0)
        part.Extension.SelectByID2("", "EDGE", 0.07, 0, 0.02, False, 4096, Nothing, 0)
        ''4096为倒角基准
        part.FeatureManager.InsertFeatureChamfer(6, 1, 0.01, pi / 4, 0, 0, 0, 0)

    End Sub

    Private Sub BunifuFlatButton14_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton14.Click
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
    Private Sub BunifuFlatButton8_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton8.Click
        'checktext()
        wt(0) = 130 / 1000 '外径  
        wt(1) = 100 / 1000 '内径
        wt(2) = 160 / 1000 '长度
        hs = 78.12 / 1000 '活塞直径
        hsg = 30 / 1000 ''活塞杆直径(接拉头螺纹直径20-25mm)
        'createwaitong()
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
        Createothers()
    End Sub
    Public Function excel()
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
                closeexcel()
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
                closeexcel()
                Exit Function
        End Select
        I = 0
        For Each sh In xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(114, 49))
            If sh.row = 114 Then
                RichTextBox1.Text = "未找到参数"
                closeexcel()
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
                        Exit For
                    End If
                End If
            End If
        Next
        GC.Collect()
    End Function
    Public Function checktext()
        If TextBox1.Text IsNot "" And TextBox2.Text IsNot "" And
            TextBox3.Text IsNot "" And TextBox4.Text IsNot "" Then
            excel()
        Else
            MsgBox("未输入值")
        End If
    End Function
    Public Function closeexcel()
        If xlBook IsNot Nothing Then
            xlBook.Close()
            xlapp.Quit()
            xlapp = Nothing
        End If
    End Function
    Public Function Createwaitong()
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
    End Function ''外筒（完成）
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
    Public Function Createothers()
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

    End Function
    Private Sub BunifuFlatButton9_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton9.Click
        swapp = CreateObject("Sldworks.Application")
        part = swapp.ActiveDoc
        parttitle = part.GetTitle
        swapp.CloseDoc(parttitle)
    End Sub
End Class
