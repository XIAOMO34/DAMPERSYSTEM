Public Class Form1
    Dim mysql As MySql.Data.MySqlClient.MySqlConnection
    ‘’‘处理窗体移动，panel2_mousedown as function ,handles panel2_mousedown as return
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr,
                                                                           ByVal wMsg As Integer,
                                                                           ByVal wParam As Integer,
                                                                           ByVal lParam As Integer) As Boolean

    Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean

    Public Const WM_SYSCOMMAND = &H112

    Public Const SC_MOVE = &HF010&

    Public Const HTCAPTION = 2
    Private Sub Panel2_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel2.MouseDown
        ReleaseCapture()

        SendMessage(Me.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub BunifuFlatButton5_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton5.Click
        Me.Hide()
    End Sub


    Private Sub BunifuFlatButton4_Click_1(sender As Object, e As EventArgs) Handles BunifuFlatButton4.Click
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub BunifuFlatButton6_Click_2(sender As Object, e As EventArgs) Handles BunifuFlatButton6.Click
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
        Dim p As Process() = Process.GetProcessesByName("MATLAB")
        For Each pr As Process In p
            pr.Kill()
        Next
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

    Private Sub BunifuFlatButton17_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton17.Click
        Dim swapp As SldWorks.SldWorks
        Dim part As SldWorks.ModelDoc2
        Dim ZhuView As Object
        Dim ZhuoView As Object
        Dim FuView As Object
        Dim myDisplayDim As Object
        swapp = CreateObject("sldworks.application")
        swapp.CloseAllDocuments(True)
        swapp.Visible = True
        part = swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\gb_a4.drwdot", 12, 0, 0)
        ''''通过模板创建新工程图
        ZhuView = part.CreateDrawViewFromModelView3("C:\Users\LJX\Desktop\零件1.SLDPRT", "*前视", 0.08, 0.16, 0)
        ''''通过指定零件创建某视图的工程图
        part.Extension.SelectByID2("", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) '工程图的类型为DRAWINGVIEW
        ZhuoView = part.CreateUnfoldedViewat3(0.2, 0.16, 0, False)
        '通过已有视图创建新视图,false表示对齐，true不对齐
        part.Extension.SelectByID2("工程图视图1", "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
        FuView = part.CreateUnfoldedViewAt3(0.08, 0.1, 0, False)
        part.ClearSelection2(True)
        part.Extension.SelectByRay(0.0835202743534853, 0.154815383481483, 250, 0, 0, -1, 0.00105527915159568, 1, False, 0, 0)
        myDisplayDim = part.AddDimension2(0.12, 0.17, 0)
    End Sub

    Private Sub BunifuFlatButton7_Click(sender As Object, e As EventArgs) Handles BunifuFlatButton7.Click
        Dim mysql As New MySql.Data.MySqlClient.MySqlConnection
        mysql.ConnectionString =
            "server=localhost;userid=root;password=Xiaomo567;database=bjpowernode"
        Try
            mysql.Open()
            MessageBox.Show("成功")
            mysql.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            mysql.Dispose()

        End Try

    End Sub
End Class
