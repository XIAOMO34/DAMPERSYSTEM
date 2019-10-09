Public Class duantou
    '模块级变量
    Dim mysqlconnect As MySql.Data.MySqlClient.MySqlConnection ''定义mysql连接
    Dim mycommand As MySql.Data.MySqlClient.MySqlCommand ''定义mysql命令
    Dim reader As MySql.Data.MySqlClient.MySqlDataReader ''定义数据流
    Dim query As String ''定义命令流
    Dim rows As Double ''定义表格列数（即参数数目）
    Dim a() As Double ''参数
    Dim swapp As SldWorks.SldWorks ''声明变量
    Dim part As SldWorks.ModelDoc2
    Dim sketch As SldWorks.SketchManager
    Dim circle1 As SldWorks.SketchSegment
    Dim feature As SldWorks.FeatureManager
    Dim lashen1 As Object
    Dim line1 As SldWorks.SketchSegment
    Dim dis As SldWorks.DisplayDimension
    Dim dimension As SldWorks.Dimension
    Dim line2 As Object
    Dim circle As Object
    Dim kong As SldWorks.Feature ''注意命名空间
    Dim xlapp As Microsoft.Office.Interop.Excel.Application ''引用Microsoft excel和Microsoft office类型库
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet ''然后创建对象
    Dim pi As Double = 3.1415926535898
    ''移动窗口代码
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr,
                                                                           ByVal wMsg As Integer,
                                                                           ByVal wParam As Integer,
                                                                           ByVal lParam As Integer) As Boolean
    Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean
    Public Const WM_SYSCOMMAND = &H112
    Public Const SC_MOVE = &HF010&
    Public Const HTCAPTION = 2
    Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles Panel1.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
    End Sub



    ''本地EXCEL调用
    Private Sub BunifuFlatButton4_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        If OpenFileDialog1.FileName = "OpenFileDialog1" Then
            MsgBox("未选择文件！")
        Else
            Useexcel()
            Createduantou()
        End If
    End Sub
    ''参数表显示
    Public Sub BunifuFlatButton1_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        OpenFileDialog1.Filter = "工作簿（*.xlsx）|*.xlsx" ''文件筛选器，只选择xlsx文件
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "OpenFileDialog1" Then
            TextBox1.Text = "文件已选择：" & OpenFileDialog1.FileName
        End If
    End Sub
    ''关闭窗口
    Private Sub BunifuImageButton1_Click(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Close()
        home.Show()
    End Sub
    ''网络数据库数据调用
    Private Sub BunifuFlatButton3_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton3.Click
        ''sql数据库
        mysqlconnect = New MySql.Data.MySqlClient.MySqlConnection ''定义连接字符串
        mysqlconnect.ConnectionString =
            "server=52.76.27.242;userid=sql12306337;password=mCA9M9cAsb;database=sql12306337" ''登录命令
        Try ''异常处理,给出弹窗提示并且暂停
            mysqlconnect.Open()
            MessageBox.Show("连接服务器成功")
            Readdata() ''读取数据库数据并且赋值，关键函数
            mysqlconnect.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mysqlconnect.Dispose() ''中断sql
        End Try
    End Sub
    ''读取网络数据库数据函数，把第二行参数赋值给a()数组，长度为字段个数，即参数个数
    Public Function Readdata() As Integer
        query = "select count(*) from information_schema.COLUMNS where table_name='damper';"
        ''读取列数，赋值给数组
        mycommand = New MySql.Data.MySqlClient.MySqlCommand(query, mysqlconnect)
        reader = mycommand.ExecuteReader ''执行sql语句
        While reader.Read
            rows = reader.GetDouble(reader.GetOrdinal("count(*)"))
        End While
        reader.Close()
        query = "select * from sql12306337.damper" ''sql语言，读取表格数据，给数据流reader
        mycommand = New MySql.Data.MySqlClient.MySqlCommand(query, mysqlconnect)
        reader = mycommand.ExecuteReader
        While reader.Read
            ReDim a(rows - 1) ''定义数组，不是单独的数，长度为rows的数组
            Dim b As Integer
            For b = 0 To rows - 1
                a(b) = reader.GetDouble("pa" & (b + 1))
                MessageBox.Show(a(b))
            Next
        End While
        Readdata = 0
    End Function
    ''建模关键函数
    Public Function Createduantou() As Integer
        MsgBox("数据已经读取，准备建立模型!")
        ''创建进程可视化
        swapp = CreateObject("Sldworks.Application")
        swapp.CloseAllDocuments(True)
        ''创建新零件
        part = swapp.NewDocument("C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2018\templates\gb_part.prtdot", 0, 0, 0)
        part = swapp.ActiveDoc
        swapp.Visible = True
        MsgBox("调用SOLIDWORKS进程成功!")
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
        MsgBox("建模成功！")
        Createduantou = 0
    End Function
    Public Function Useexcel() As Integer ''调用EXCEL参数函数
        xlapp = CreateObject("Excel.Application") ''创建EXCEL对象
        ''打开已经存在的EXCEL工件簿文件
        xlBook = xlapp.Workbooks.Open(OpenFileDialog1.FileName)
        xlSheet = xlBook.Worksheets("sheet1")
        ''运用数组存数据方便实用（都是double）
        Dim a() As Double
        ReDim a(1)
        Dim i As Integer
        For i = 0 To a.Length - 1
            a(i) = xlSheet.Cells(2, i + 1).value
        Next
        ''关闭excel进程，释放内存
        Dim p As Process() = Process.GetProcessesByName("EXCEL")
        For Each pr As Process In p
            pr.Kill()
        Next
        Useexcel = 0
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

    Private Sub BunifuFlatButton5_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton5.Click
        Createlatou()
    End Sub
End Class