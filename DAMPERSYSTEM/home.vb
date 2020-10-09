Imports SldWorks
Imports Microsoft.Office.Interop.Excel
Imports MySql.Data.MySqlClient

Public Class home ''kenan
    Dim mysql As MySqlConnection
    Dim swapp As SldWorks.SldWorks
    Dim part As ModelDoc2
    Dim asm As AssemblyDoc
    Dim mysqlconnect As MySqlConnection ''定义mysql连接
    Dim mycommand As MySqlCommand ''定义mysql命令
    Dim reader As MySqlDataReader ''定义数据流

    ''处理窗体移动，panel2_mousedown as function ,handles panel2_mousedown as return
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As IntPtr,
         ByVal wMsg As Integer,
         ByVal wParam As Integer,
         ByVal lParam As Integer) As Boolean
    Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean

    Public Const WM_SYSCOMMAND = &H112

    Public Const SC_MOVE = &HF010&

    Public Const HTCAPTION = 2

    Private Sub Panel2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) _
        Handles Panel2.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'mysqlconnect = New MySqlConnection ''定义连接字符串
        'mysqlconnect.ConnectionString =
        '    "server=52.76.27.242;userid=sql12307948;password=W38GxxRxLI;database=sql12307948" ''登录命令
        'mysqlconnect.Open()
        'If mysqlconnect.State = 1 Then
        '    MessageBox.Show("1")
        'End If
        ''Try ''异常处理,给出弹窗提示并且暂停
        ''    mysqlconnect.Open()
        ''    Label2.Text = "服务器状态：已连接"
        ''    mysqlconnect.Close()
        ''Catch ex As Exception
        ''    Label2.Text = "服务器状态：未连接"
        ''Finally
        ''    mysqlconnect.Dispose() ''中断sql
        ''End Try
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

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

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub
    Private Sub BunifuFlatButton1_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton1.Click
        partpanel.Visible = True
        assemblepanel.Visible = False
        morepanel.Visible = False
        setdamper.Visible = False
        drawingpanel.Visible = False
    End Sub

    Private Sub BunifuImageButton1_Click_2(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
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
        partsub.Label1.Text = "小头端盖生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\xtdg.JPG")
        Me.Hide()
        partsub.Show()
    End Sub

    Private Sub Buttonneitong_Click(sender As Object, e As EventArgs) Handles Buttonneitong.Click
        partsub.Label1.Text = "大头端盖生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\dtdg1.JPG")
        Me.Hide()
        partsub.Show()
    End Sub

    Private Sub Buttonzunitong_Click(sender As Object, e As EventArgs) Handles Buttonzunitong.Click
        partsub.Label1.Text = "外筒生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\wt方形.png")
        Me.Hide()
        partsub.Show()

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub

    Private Sub assemblepanel_Paint(sender As Object, e As PaintEventArgs) Handles assemblepanel.Paint

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

    Private Sub partpanel_Paint(sender As Object, e As PaintEventArgs) Handles partpanel.Paint

    End Sub

    Private Sub BunifuFlatButton4_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        partsub.Label1.Text = "活塞杆生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\hsg.JPG")
        Me.Hide()
        partsub.Show()
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

    End Sub

    Private Sub BunifuFlatButton7_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton7.Click
        partsub.Label1.Text = "拉头生成子程序"
        partsub.PictureBox3.Load("D:\POST-GRA\研究生大论文\论文素材\图片\lt.JPG")
        Me.Hide()
        partsub.Show()
    End Sub
End Class
