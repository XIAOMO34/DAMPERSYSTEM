Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop.Excel
Imports SldWorks ''命名空间 
Public Class partsub
    '模块级变量
    Public parttype As Single
    Dim mysqlconnect As MySqlConnection ''定义mysql连接
    Dim mycommand As MySqlCommand ''定义mysql命令
    Dim reader As MySqlDataReader ''定义数据流
    Dim query As String ''定义命令流
    Dim rows As Double ''定义表格列数（即参数数目）
    Dim a() As Double ''参数
    Dim swapp As SldWorks.SldWorks ''声明变量
    Dim part As ModelDoc2
    Dim sketch As SketchManager
    Dim circle1 As SketchSegment
    Dim feature As FeatureManager
    Dim lashen1 As Object
    Dim line1 As SketchSegment
    Dim dis As DisplayDimension
    Dim dimension As Dimension
    Dim line2 As Object
    Dim circle As Object
    Dim kong As Feature ''注意命名空间
    Public pi As Double = 3.1415926535898
    ''窗口1关闭时主窗口打开
    Private Sub partsub_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        home.Show()
    End Sub

    ''关闭窗口
    ''建模关键函数
    Public Function Createduantou() As Integer
        ''创建进程可视化
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        ''创建新零件
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        swapp.Visible = True
        sketch = part.SketchManager
        ''插入草图
        part.InsertSketch()
        ''创建圆
        circle1 = sketch.CreateCircle(0, 0, 0, 0.01, 0, 0)
        feature = part.FeatureManager
        ''拉伸圆
        lashen1 = feature.FeatureExtrusion3(True, True, False, 0, 0,
                                                0.01, 0, False, False, False,
                                                False, 0, 0, True, True,
                                                False, False, True, False, False,
                                                0, 0, 0)
        part.Extension.SelectByID2("", "face", 0, 0, 0.01, False, 0, Nothing, 0)
        part.InsertSketch()
        ''part.ClearSelection2(True)
        ''创建直线
        line1 = sketch.CreateLine(0.006, 0.008, 0, 0.006, -0.008, 0)
        ''修改尺寸（尺寸驱动法）重要,需要先用getdimension ,再用setsystemvalue方法改变尺寸
        dis = part.AddDimension2(0.015, 0, 0)
        dimension = dis.GetDimension2(0)
        dimension.SetSystemValue3(0.01, -1, "")
        line2 = sketch.CreateLine(0.006, 0.008, 0, 0.006, -0.008, 0)
        circle = sketch.CreateCircle(0, 0, 0, 0.01, 0, 0)
        part.ClearSelection2(True)
        ''swapp.SendMsgToUser(sb)'获取对象名字，用来选取对象
        ''裁剪不需要的部分（此处虽然选的是整个圆，但仍可模拟鼠标点取的操作）
        ''就算选的是整个圆，仍然可以当做一部分裁剪掉
        part.Extension.SelectByID2("", "arc", 0.01, 0, 0.01, False, 2, Nothing, 0)
        part.SketchManager.SketchTrim(4, 0, 0, 0)
        part.Extension.SelectByID2("", "arc", -0.01, 0, 0.01, False, 2, Nothing, 0)
        part.SketchManager.SketchTrim(4, 0, 0, 0)
        ''也可以知道对象名字后用“sketchsegment”
        part.ClearSelection2(True)
        ''再次拉伸
        part.Extension.SelectByID2("", "face", 0.0065, 0, 0, False, 0, Nothing, 0)
        feature.FeatureExtrusion3(True, True, False, 0, 0,
                                       0.01, 0, False, False, False,
                                       False, 0, 0, True, True,
                                       False, False, True, False, False,
                                       0, 0, 0)
        part.ClearSelection2(True)
        part.Extension.SelectByID2("", "face", 0, 0, 0.01, False, 0, Nothing, 0)
        part.InsertSketch()
        sketch.CreateCenterLine(0, 0.01, 0, 0, -0.01, 0)
        part.ClearSelection2(True)
        part.InsertSketch()
        part.Extension.SelectByID2("", "face", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT",
                                       0, 0.00527234725331027, 0, True, 1, Nothing, 0)
        ''创建基准面，为镜像做准备
        Dim myRefPlane As Object
        myRefPlane = part.FeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0)
        part.ClearSelection2(True)
        part.Extension.SelectByID2("凸台-拉伸2", "BODYFEATURE", 0, 0, 0, False, 1, Nothing, 0)
        part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 2, Nothing, 0)
        ''镜像特征（标记分别用1，2，见API文件）
        part.FeatureManager.InsertMirrorFeature2(False, False, True, True, 0)
        part.Extension.SelectByID2("", "FACE", 0.006, 0, 0.011, False, 0, Nothing, 0)
        part.InsertSketch()
        sketch.CreateCircle(0.015, 0, 0, 0.015, 0.002, 0)
        part.InsertSketch()
        ''创建切除特征
        feature.FeatureCut4(False, False, False, 9, 1,
                                0.01, 0.01, False, False, False,
                                False, 0.0174532925199433, 0.0174532925199433, False, False,
                                False, False, False, True, True,
                                True, True, False, 0, 0,
                                False, False)
        part.Extension.SelectByID2("", "FACE", 0, 0, 0, False, 0, Nothing, 0)
        ''创建孔
        Dim swWzdHole As SldWorks.WizardHoleFeatureData2
        swWzdHole = feature.CreateDefinition(SwConst.swFeatureNameID_e.swFmHoleWzd)
        swWzdHole.InitializeHole(SwConst.swWzdGeneralHoleTypes_e.swWzdCounterBore,
                                     SwConst.swWzdHoleStandards_e.swStandardISO,
                                     SwConst.swWzdHoleStandardFastenerTypes_e.swStandardISOHexBolt,
                                     "M5",
                                     SwConst.swEndConditions_e.swEndCondBlind)
        part.Extension.SelectByID2("", "FACE", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("", "FACE", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("", "FACE", 0, 0, 0, False, 0, Nothing, 0)
        kong = feature.CreateFeature(swWzdHole)
        ''重合孔与原点，定义正确位置
        part.Extension.SelectByID2("草图6", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        part.InsertSketch()
        ''先建立孔，再利用约束找位置
        part.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        part.Extension.SelectByID2("Point1@原点", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, 0)
        part.SketchAddConstraints("sgCOINCIDENT")
        part.ClearSelection2(True)
        part.SketchManager.InsertSketch(True)
        Createduantou = 0
    End Function
    Public Function createwaitong()
        ''创建进程可视化
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        ''创建新零件
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        swapp.Visible = True
        part = swapp.ActiveDoc
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.SketchManager.InsertSketch(True)
        part.ClearSelection2(True)
        part.SketchManager.CreateCircle(0#, 0#, 0#, 0.06, 0, 0#)
        part.SketchManager.CreateCircle(0#, 0#, 0#, 0.08, 0, 0#)
        part.FeatureManager.FeatureExtrusion3(True, True, False, 0, 0,
                                                0.2, 0, False, False, False,
                                                False, 0, 0, True, True,
                                                False, False, True, False, False,
                                                0, 0, 0)
        part.ClearSelection2(True)
        part.Extension.SelectByID2("", "FACE", 0.07, 0, 0.2, False, 0, Nothing, 0)
        part.SketchManager.InsertSketch(True)
        part.SketchManager.CreateCircle(0, 0, 0.2, 0.07, 0, 0.2)
        part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.01, 0.01, False,
                                        False, False, False, 0, 0,
                                        False, False, False, False, False, True, True, True, True, False,
                                         0, 0, False, False)
        part.ShowNamedView2("*前视", 1)
        part.Extension.SelectByID2("", "", 0.07, 0, 0.2, False, 0, Nothing, 0)
        part.FeatureManager.InsertCosmeticThread2(1, 0.14, 0, "140")
        part.ShowNamedView2("", 7)
        part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.FeatureManager.InsertRefPlane(8, 0.08, 0, 0, 0, 0)
        'part.Extension.SelectByID2("", "PLANE", 0.08, 0, 0, False, 0, Nothing, 0)
        'part.SketchManager.InsertSketch(True)
        'part.SketchManager.CreateCircleByRadius(-0.15, 0, 0, 0.01) ''草图只有二维，Z无意义
        '''''''''''''''''''''''''''''''''''
        Dim swWzdHole As WizardHoleFeatureData2
        swWzdHole = part.FeatureManager.CreateDefinition(SwConst.swFeatureNameID_e.swFmHoleWzd)
        part.Extension.SelectByID2("", "FACE", 0.08, 0, 0.15, False, 0, Nothing, 0)
        part.FeatureManager.HoleWizard5(4, 1, 42, "M10x1.0", 2, 0.009, 0.02, 0.02, 0, 0, 0, 0, 0, 0, 2, 0, 0, -1, -1, -1, "", False, True, True, True, True, False)
        part.Extension.SelectByID2("凸台-拉伸1", "BODYFEATURE", 0, 0, 0, False, 1, Nothing, 0) ''1是镜像特征
        part.Extension.SelectByID2("M10x1.0 螺纹孔1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        part.Extension.SelectByID2("切除-拉伸1"， "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 2, Nothing, 0) ''2是镜像基准
        part.FeatureManager.InsertMirrorFeature(False, False, False, False)
    End Function
    Public Function createdatouduangai()    ''最大直径为活塞杆直径，最小为加上盖头
        swapp = CreateObject("Sldworks.Application")
        'swapp.CloseAllDocuments(True)
        ''创建新零件
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Cylinider head.SLDPRT",
                              1, 0, "", 0, 0)
        part = swapp.ActiveDoc
        part.ActiveView.FrameState = 1
        part.Parameter("D10@Sketch1").SYSTEMVALUE = 0.047 ''修改最大值
        part.EditRebuild3()
    End Function
    Public Function createxiaotouduangai()
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        ''创建新零件
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Cylinider head.SLDPRT",
                              1, 0, "", 0, 0)
        part = swapp.ActiveDoc
    End Function
    Public Function createhuosaigan()
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        ''创建新零件
        swapp.Visible = True
        part = swapp.OpenDoc6("D:\POST-GRA\研究生大论文\零件库\500KN液压抗震阻尼器\Rod.SLDPRT",
                              1, 0, "", 0, 0)
        part = swapp.ActiveDoc
        part.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        part.EditSketch() ''活塞杆直径
        part.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        part.EditSketch()
        part.Parameter("D2@Sketch1").systemvalue = 0.042
        part.EditRebuild3()
    End Function
    Public Function Createlatou()
        swapp = CreateObject("sldworks.Application")
        swapp.CloseAllDocuments(True)
        swapp.Visible = True
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        sketch = part.SketchManager
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.InsertSketch()
        circle1 = sketch.CreateCircle(0, 0, 0, 0, 0.03, 0)
        feature = part.FeatureManager
        lashen1 = feature.FeatureExtrusion3(True, True, False, 0, 0,
                                            0.1, 0, False, False, False,
                                            False, 0, 0, True, True,
                                            False, False, True, False, False,
                                            0, 0, 0)
        part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.FeatureManager.InsertRefPlane(264, 0.02, 0, 0, 0, 0)
        part.Extension.SelectByID2(""， "PLANE"， 0， 0, -0.02， False， 0， Nothing, 0)
        part.InsertSketch()
        sketch.CreateCircle(0, 0, 0.02, 0, 0.05, 0.02)
        lashen1 = feature.FeatureExtrusion3(True, True, True, 0, 0,
                                            0.16, 0, False, False, False,
                                            False, 0, 0, True, True,
                                            False, False, True, False, False,
                                            0, 0, 0)
        part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        part.ShowNamedView2（"*右视", 4）
        part.ViewZoomtofit2()
        part.InsertSketch()
        part.SketchManager.CreateLine(-0#, 0#, 0#, -0#, 0.03, 0#)
        part.SketchManager.CreateLine(-0#, 0.03, 0#, 0.02, 0.05, 0#)
        part.SketchManager.CreateLine(0.02, 0.05, 0#, 0.02, 0#, 0#)
        part.SketchManager.CreateLine(0.02, 0#, 0#, -0#, 0#, 0#)
        part.SketchManager.CreateCenterLine(-0#, 0#, 0#, -0.1, 0#, 0#)
        part.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, False, 16, Nothing, 0)
        part.FeatureManager.FeatureRevolve2(True, True, False, False, False,
                                            False, 0, 0, 2 * (pi), 0,
                                            False, False, 0.01, 0.01, 0,
                                            0, 0, True, True, True)
        ''pi乘2为旋转角度，应当为double精度，即旋转360度
        part.ShowNamedView2("", 2)
        part.Extension.SelectByID2("", "FACE", 0, 0, -0.18, False, 0, Nothing, 0)
        part.InsertSketch()
        sketch.CreatePoint(0, 0, 0)
        part.Extension.SelectByID2("", "FACE", 0, 0, -0.18, False, 0, Nothing, 0)
        'swWzdhole = feature.CreateDefinition(SwConst.swFeatureNameID_e.swFmHoleWzd)
        'swWzdhole.InitializeHole(2, 13, 357, "M64X3.0", 0)
        'part.Extension.SelectByID2("", "FACE", 0, 0, -0.18, False, 0, Nothing, 0)
        'Hole = feature.CreateFeature(swWzdhole)
        feature.HoleWizard5(2, 13, 357, "M64X3.0", 0,
                            0.064, 0.07, -1, 1, pi,
                            0, 0, 0, 0, 0,
                            -1, -1, -1, -1, -1,
                            "", False, True, True, True,
                            True, False)
        part.ShowNamedView2("", 2)
        part.InsertSketch()
        part.Extension.SelectByID2("", "FACE", 0, 0, -0.11, False, 0, Nothing, 0)
        part.InsertSketch()
        sketch.CreateCircle(0, 0, 0, 0, 0.02, 0)
        feature.FeatureCut4(True, False, False, 0, 0,
                            0.095, 0.01, False, False, False,
                            False, 0, 0, False, False,
                            False, False, False, True, True,
                            True, True, False, 0, 0,
                            False, False)
        'Me.Close()
        Createlatou = 0
    End Function



    Private Sub Button2_Click(sender As Object, e As EventArgs)
        'If OpenFileDialog1.FileName = "OpenFileDialog1" Then
        '    MsgBox("未选择文件！")
        'Else
        'Useexcel()
        Me.WindowState = 1

        Select Case parttype''零件类型
            Case 1
                'Createduantou()
                home.Createwaitong()
            Case 2
                'createwaitong()
                home.Createrod()
            Case 3
                'createdatouduangai()
                home.Createpiston()
        End Select
        'End If
        Me.WindowState = 0
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'If OpenFileDialog1.FileName = "OpenFileDialog1" Then
        '    MsgBox("未选择文件！")
        'Else
        'Useexcel()
        Me.WindowState = 1

        Select Case parttype''零件类型
            Case 1
                'Createduantou()
                home.Createwaitong()
            Case 2
                'createwaitong()
                home.Createrod()
            Case 3
                'createdatouduangai()
                home.Createpiston()
        End Select
        'End If
        Me.WindowState = 0
    End Sub
End Class
'注释段
'移动窗口代码
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr,
'ByVal wMsg As Integer,
'ByVal wParam As Integer,
'ByVal lParam As Integer) As Boolean
'Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean
'Public Const WM_SYSCOMMAND = &H112
'Public Const SC_MOVE = &HF010&
'Public Const HTCAPTION = 2
'Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
'    ReleaseCapture()
'    SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
'End Sub
''网络数据库数据调用
'Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs)
'    ''sql数据库
'    mysqlconnect = New MySql.Data.MySqlClient.MySqlConnection ''定义连接字符串
'    mysqlconnect.ConnectionString =
'            "server=52.76.27.242;userid=sql12307948;password=W38GxxRxLI;database=sql12307948" ''登录命令
'    Try ''异常处理,给出弹窗提示并且暂停
'        mysqlconnect.Open()
'        MessageBox.Show("连接服务器成功")
'        Readdata() ''读取数据库数据并且赋值，关键函数k
'        mysqlconnect.Close()
'    Catch ex As Exception
'        MessageBox.Show(ex.Message)
'    Finally
'        mysqlconnect.Dispose() ''中断sql
'    End Try
'End Sub
''EXCEL调用及关闭
'Public Function Useexcel() As Integer ''调用EXCEL参数函数
'    xlapp = CreateObject("Excel.Application") ''创建EXCEL对象
'    xlBook = xlapp.Workbooks.Open(OpenFileDialog1.FileName) ''打开已经存在的EXCEL工件簿文件
'    xlSheet = xlBook.Worksheets("sheet1")
'    Dim a() As Double ''运用数组存数据方便实用（都是double）
'    ReDim a(1)
'    Dim i As Integer
'    For i = 0 To a.Length - 1
'        a(i) = xlSheet.Cells(2, i + 1).value
'    Next
'    ''关闭excel进程，释放内存
'    Dim p As Process() = Process.GetProcessesByName("EXCEL")
'    For Each pr As Process In p
'        pr.Kill()
'    Next
'    Useexcel = 0
'End Function
''读取网络数据库数据函数，把第二行参数赋值给a()数组，长度为字段个数，即参数个数
'Public Function Readdata() As Integer
'    query = "select count(*) from information_schema.COLUMNS where table_name='damper';"
'    ''读取列数，赋值给数组
'    mycommand = New MySql.Data.MySqlClient.MySqlCommand(query, mysqlconnect)
'    reader = mycommand.ExecuteReader ''执行sql语句
'    While reader.Read
'        rows = reader.GetDouble(reader.GetOrdinal("count(*)"))
'    End While
'    reader.Close()
'    query = "select * from sql12307948.damper" ''sql语言，读取表格数据，给数据流reader
'    mycommand = New MySql.Data.MySqlClient.MySqlCommand(query, mysqlconnect)
'    reader = mycommand.ExecuteReader
'    While reader.Read
'        ReDim a(rows - 1) ''定义数组，不是单独的数，长度为rows的数组
'        Dim b As Integer
'        For b = 0 To rows - 1
'            a(b) = reader.GetDouble("pa" & (b + 1))
'            MessageBox.Show(a(b))
'        Next
'    End While
'    Readdata = 0
'End Function
'Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
'    'mysqlconnect = New MySqlConnection ''定义连接字符串
'    'mysqlconnect.ConnectionString =
'    '    "server=52.76.27.242;userid=sql12307948;password=W38GxxRxLI;database=sql12307948" ''登录命令
'    'mysqlconnect.Open()
'    'If mysqlconnect.State = 1 Then
'    '    MessageBox.Show("1")
'    'End If
'    ''Try ''异常处理,给出弹窗提示并且暂停
'    ''    mysqlconnect.Open()
'    ''    Label2.Text = "服务器状态：已连接"
'    ''    mysqlconnect.Close()
'    ''Catch ex As Exception
'    ''    Label2.Text = "服务器状态：未连接"
'    ''Finally
'    ''    mysqlconnect.Dispose() ''中断sql
'    ''End Try
'End Sub

''处理窗体移动，panel2_mousedown as function ,handles panel2_mousedown as return
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'        (ByVal hwnd As IntPtr,
'         ByVal wMsg As Integer,
'         ByVal wParam As Integer,
'         ByVal lParam As Integer) As Boolean
'Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean

'Public Const WM_SYSCOMMAND = &H112

'Public Const SC_MOVE = &HF010&

'Public Const HTCAPTION = 2

'Private Sub Panel2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)

'    ReleaseCapture()
'    SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
'End Sub
'Private Sub BunifuFlatButton15_Click(sender As Object, e As EventArgs)
'    ''创建进程可视化
'    swapp = CreateObject("Sldworks.Application")
'    ''swapp.CloseAllDocuments(True)
'    ''创建新零件
'    part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
'    part = swapp.ActiveDoc
'    swapp.Visible = True
'    part = swapp.ActiveDoc
'    part.Extension.SelectByID2("前视基准面"， "PLANE"， 0， 0, 0， False， 0， Nothing， 0)
''    part.SketchManager.InsertSketch(True)
''    part.SketchManager.CreateCircleByRadius(0, 0, 0, 0.07)
''    part.FeatureManager.FeatureExtrusion3(True, True, False, 0, 0,
''                                                0.02, 0, False, False, False,
''                                                False, 0, 0, True, True,
''                                                False, False, True, False, False,
''                                                0, 0, 0)
''    part.Extension.SelectByID2("", "EDGE", 0.07, 0, 0.02, False, 4096, Nothing, 0)
''    ''4096为倒角基准
''    part.FeatureManager.InsertFeatureChamfer(6, 1, 0.01, pi / 4, 0, 0, 0, 0)

''End Sub
'Private Sub BunifuFlatButton5_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton5.Click
'    'Me.Hide()
'    Dim featruemgr As FeatureManager
'    Dim datouduangai As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\大头端盖.SLDPRT"
'    Dim xiaotouduangai As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\小头端盖.SLDPRT"
'    Dim latou As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\拉头.SLDPRT"
'    Dim jietou As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\接头.SLDPRT"
'    Dim waitong As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\外筒.SLDPRT"
'    Dim huosaigan As String =
'            "C:\Users\LJX\Desktop\装配练习2019-10-14\活塞杆.SLDPRT"
'    Dim title As String
'    swapp = CreateObject("Sldworks.Application")
'    swapp.CloseAllDocuments(True)
'    swapp.Visible = True
'    part = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\gb_assembly.asmdot", 0,
'                                 0, 0)
'    title = part.GetTitle
'    asm = part
'    Addcomponent(waitong)
'    Addcomponent(huosaigan)
'    Addcomponent(datouduangai)
'    Addcomponent(xiaotouduangai)
'    swapp.OpenDoc6(latou, 1, 32, "", 2, 2)
'    asm.AddComponent5(latou, 0, "", False, "", 1, 0, 0)
'    swapp.CloseDoc(latou)
'    Addcomponent(jietou)
'    part.Extension.SelectByID2("右视基准面@大头端盖-1@" & title, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
'    part.Extension.SelectByID2("右视基准面@外筒-1@" & title, "PLANE", 0, 0, 0, True, 0, Nothing, 0)
'    asm.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0,
'                     0, 0, False, False, 0, 0)
'    part.ShowNamedView2("*左视", 3)
'    part.Extension.SelectByID2("右视基准面@接头-1@" & title, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
'    part.Extension.SelectByID2("", "FACE", -0.153, 0.054, 0, True, 0, Nothing, 0)
'    asm.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0,
'                     0, 0, False, False, 0, 0)
'    part.Extension.SelectByID2("", "FACE", 0.153, 0.05, 0, False, 0, Nothing, 0)
'    part.ShowNamedView2("*左视", 3)
'    part.Extension.SelectByID2("", "FACE", 1 - 0.28, 0.035, 0, True, 0, Nothing, 0)
'    part.ShowNamedView2("*", 7)
'    asm.AddMate5(0, 1, False, 0.001, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, False, False, 0, 1)
'End Sub
'Function Addcomponent(a As String)
'    swapp.OpenDoc6(a, 1, 32, "", 2, 2)
'    asm.AddComponent5(a, 0, "", False, "", 0, 0, 0)
'    swapp.CloseDoc(a)
'    Return 0
'End Function
'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
'    OpenFileDialog1.Filter = "工作簿（*.xlsx）|*.xlsx" ''文件筛选器，只选择xlsx文件
'    OpenFileDialog1.ShowDialog()
'    If OpenFileDialog1.FileName <> "OpenFileDialog1" Then
'        TextBox1.Text = "文件已选择：" & OpenFileDialog1.FileName
'    End If
'End Sub