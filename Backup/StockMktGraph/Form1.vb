Imports System.Threading

Public Class Form1

    Dim x_mm, y_mm As Int16     'cordinats of the line
    Dim myPen = New Pen(Color.Blue) 'graph pen color
    Dim gridPen = New Pen(Color.LightGray) 'grid pen color
    Dim graphics As Graphics

    Structure stockData
        Dim sdate As Date
        Dim open As Single
        Dim high As Single
        Dim low As Single
        Dim close As Single
        Dim vol As Integer
    End Structure
    Structure sdSize
        Dim symbol As String        ' the stock symbol
        Dim n As Integer            ' the # of elements in the stockData array that is
        '                           '   returned to the calling program by getAH
    End Structure
    Public Sub getAH(ByRef sd() As stockData, ByRef sd1 As sdSize)
        Dim inpfilenum As Integer, inpfilename As String
        Dim filebase As String
        Dim stg As String
        Dim k As Integer
        Dim alocal() As String
        Dim localArray(300) As stockData
        '
        ' NOTE TO CODER: for your testing, you can set your own filebase and filename
        '                what I have here is what works for me
        '
        'filebase = "D:\VisualStudioProjects\StockMktGraph\StockMktGraph\"
        filebase = "..\"
        inpfilenum = FreeFile()
        inpfilename = filebase & "AH total.txt"
        FileOpen(inpfilenum, inpfilename, OpenMode.Input, OpenAccess.Read)
        '
        ' read the first line and get the symbol and size
        '
        stg = LineInput(inpfilenum)
        alocal = Split(stg, ",")
        sd1.symbol = alocal(0)
        sd1.n = CInt(alocal(1))
        '
        ' read the rest of the lines and put the data into sd
        '
        Thread.CurrentThread.CurrentCulture = New Globalization.CultureInfo("en-US")
        For ix = 0 To sd1.n - 1
            '
            ' get a line from the file and split it into the various parts
            '
            stg = LineInput(inpfilenum)
            alocal = Split(stg, ",")
            '
            ' convert the date string (which is in a format of MY preference) into a date variable
            '
            stg = Mid(alocal(0), 5, 2) & "/" & Mid(alocal(0), 7, 2) & "/" & Microsoft.VisualBasic.Left(alocal(0), 4)
            '
            ' get the rest of this line into sd
            '

            sd(ix).sdate = CDate(stg)
            sd(ix).open = CSng(alocal(1))
            sd(ix).high = CSng(alocal(2))
            sd(ix).low = CSng(alocal(3))
            sd(ix).close = CSng(alocal(4))
            sd(ix).vol = CInt(alocal(5))
        Next
        FileClose(inpfilenum)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove

        'Copy mouse move position
        x_mm = e.X
        y_mm = e.Y

        'Init PictureBox
        Dim bmp01 As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        PictureBox1.Image = bmp01
        GC.SuppressFinalize(bmp01)
        GC.Collect()

        'create array of stockData
        Dim sd() As stockData = New stockData(300) {}
        Dim sd1 As sdSize

        'call getAH for sd updation
        getAH(sd, sd1)

        'Graphics object initialization
        Dim bmp As New Bitmap(PictureBox1.Image)
        graphics = graphics.FromImage(bmp)
        PictureBox1.Image = bmp

        'x,y coordinates of graph
        Dim x As Single, y As Single

        'erase lines
        graphics.Clear(Color.White)

        'initial x,y position
        x = 0
        y = 0
        'interval for scaling
        Dim interval As Single

        'Vertical grid drawing
        For cnt = 1 To 25
            graphics.DrawLine(gridPen, 0, cnt * 25, 1008, cnt * 25)
        Next

        'Horigontal grid drawing
        For cnt = 1 To 12
            graphics.DrawLine(gridPen, cnt * 84, 0, cnt * 84, 500)
        Next
        For ix = 0 To sd1.n - 1
            'read sd().open values
            x = sd(ix).open * 10
            y = sd(ix + 1).open * 10

            'Scaling calcution
            x = CInt(x)
            y = CInt(y)
            x = 350 - x
            y = 350 - y
            x = x + (x * 0.7)
            y = y + (y * 0.7)
            x = x / 0.7
            y = y / 0.7

            'draw graph
            graphics.DrawLine(myPen, interval, x, interval + 4, y)
            interval = interval + 4
        Next

        Dim myPen1 = New Pen(Color.Red)

        'Draw red line cursor
        graphics.DrawLine(myPen1, x_mm, 0, x_mm, 500)

        'Update lables with sd() data
        Label1.Text = "Stock Symbol Date = " & sd(x_mm / 4).sdate
        Label2.Text = "Open = " & sd(x_mm / 4).open
        Label3.Text = "High = " & sd(x_mm / 4).high
        Label4.Text = "Low =  " & sd(x_mm / 4).low
        Label5.Text = "Close =  " & sd(x_mm / 4).close
        Label6.Text = "Volume = " & sd(x_mm / 4).vol
        Refresh()
        '
        'Release resources
        GC.Collect()
    End Sub
End Class
