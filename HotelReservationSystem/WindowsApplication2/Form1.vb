Imports System.Data.OleDb
Imports System.Configuration


Imports System.IO
Public Class Form1
    Dim SqlDa As OleDbDataAdapter
    Dim sqlcommand As OleDbCommand

    Dim DatSt As New DataSet
    Dim SqlCn As New OleDbConnection
    Dim SQLstr As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (DateTimePicker1.Value < DateTimePicker2.Value) Then
            SQLstr = ""
            SQLstr = SelectAvailableRooms(DateTimePicker1.Value.Date, DateTimePicker2.Value.Date, False)

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "Search")
            DataGridView1.DataSource = DatSt.Tables("Search")
            DatSt.Tables.Remove("Search")

            SQLstr = SelectAvailableRooms(DateTimePicker1.Value.Date, DateTimePicker2.Value.Date, True)
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "CountSearch")
            DataGridView2.DataSource = DatSt.Tables("CountSearch")
            DatSt.Tables.Remove("CountSearch")
        Else
            MsgBox("CheckInDate must greater than CheckOutDate")
            Return

        End If
        



    End Sub

   
   

 
 
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        ' Display the date as "Mon 26 Feb 2001".
        DateTimePicker1.CustomFormat = "yyyy/MM/dd"
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        ' Display the date as "Mon 26 Feb 2001".
        DateTimePicker2.CustomFormat = "yyyy/MM/dd"
        SqlCn = New OleDbConnection()
        SqlCn.ConnectionString = ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString()

        SqlCn.Open()
        Dim tmp As String
        If SqlCn.State = ConnectionState.Open Then
            tmp = "Connection opened successfully"
        Else
            tmp = "Connection could not be established"
        End If

        SqlCn.Close()
        Label4.Text = "$"
        Label5.Text = ""
        Label6.Text = ""

        ''initailizing value for Ranch service(Combobox1)
        SQLstr = ""
        SQLstr = "SELECT * FROM RanchService"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "RanchType")
        ComboBox1.DataSource = DatSt.Tables("RanchType")
        ComboBox1.DisplayMember = "Service Type"
        ComboBox1.ValueMember = "Service ID"

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        

        DatSt.Tables.Remove("RanchType")

        SQLstr = ""
        SQLstr = " SELECT COUNT(RoomID) AS Total,[Room Type] FROM Roomtable group by [Room Type]"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "RoomType")
        DataGridView2.DataSource = DatSt.Tables("RoomType")
        DatSt.Tables.Remove("RoomType")
        Label61.Visible = False

        ComboBox5.Visible = False



    End Sub

   
   
  
    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged

        Dim RoomType As String = ""
        Dim tmp As String() = DateCheck(DateTimePicker1.Value.Date, DateTimePicker2.Value.Date).Split(",")
        If (DataGridView1.CurrentRow.Selected) Then
            RoomType = DataGridView1.SelectedRows(0).Cells("Room Type").Value

            Label4.Text = DateCheck(DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)


            'tmp(0) is how many days of peakweekday 
            'tmp(1) is peakweekend 
            'tmp (2) is notpeakweekday
            'tmp(3) is notpeakweekend
            ' Label4.Text = tmp(0) & "_" & tmp(1) & "_" & tmp(2) & "_" & tmp(3)
            SQLstr = ""

            SQLstr = SelectRoomsRATE(RoomType, tmp(3), tmp(2), tmp(1), tmp(0))

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)

            SqlDa.Fill(DatSt, "RATE")
            Label4.Text = "$" & DatSt.Tables("RATE").Rows(0)(0).ToString

            Label5.Text = tmp(3) & " Day(s) in Weekend " & tmp(2) & " Day(s) in Weekday /n "
            Label6.Text = tmp(1) & " Day(s) in Weekend(Peak) " & tmp(0) & " Day(s) in Weekday(Peak)"
            DatSt.Tables.Remove("RATE")
            SQLstr = ""
        End If
        Dim CountGroupRooms As Integer = 0

        If DataGridView1.SelectedRows.Count > 1 Then
            For i As Integer = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Rows(i).Cells(2).Value.ToString = "Single" Then
                    If DataGridView1.Rows(i).Selected = True Then
                        RoomType = "Single"
                        SQLstr = ""

                        SQLstr = SelectRoomsRATE(RoomType, tmp(3), tmp(2), tmp(1), tmp(0))

                        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)

                        SqlDa.Fill(DatSt, "RATE")
                        CountGroupRooms = CountGroupRooms + Int(DatSt.Tables("RATE").Rows(0)(0).ToString())
                        DatSt.Tables.Remove("RATE")
                    End If
                End If
                If DataGridView1.Rows(i).Cells(2).Value.ToString = "Double" Then
                    If DataGridView1.Rows(i).Selected = True Then
                        RoomType = "Double"
                        SQLstr = ""

                        SQLstr = SelectRoomsRATE(RoomType, tmp(3), tmp(2), tmp(1), tmp(0))

                        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)

                        SqlDa.Fill(DatSt, "RATE")
                        CountGroupRooms = CountGroupRooms + Int(DatSt.Tables("RATE").Rows(0)(0).ToString())
                        DatSt.Tables.Remove("RATE")
                    End If


                End If
                If DataGridView1.Rows(i).Cells(2).Value.ToString = "Quadruple" Then
                    If DataGridView1.Rows(i).Selected = True Then
                        RoomType = "Quadruple"
                        SQLstr = ""

                        SQLstr = SelectRoomsRATE(RoomType, tmp(3), tmp(2), tmp(1), tmp(0))

                        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)

                        SqlDa.Fill(DatSt, "RATE")
                        CountGroupRooms = CountGroupRooms + Int(DatSt.Tables("RATE").Rows(0)(0).ToString())
                        DatSt.Tables.Remove("RATE")
                    End If

                End If
            Next
            Label4.Text = "$" & CountGroupRooms.ToString & "-for " & DataGridView1.SelectedRows.Count & " rooms"
            Label61.Text = CountGroupRooms.ToString
        End If


    End Sub
    Public Shared Function DateCheck(ByVal StartDate As Date, ByVal EndDate As Date) As String
        Dim count As Integer
        Dim PeakWeekday As Integer
        Dim PeakWeekend As Integer
        Dim NotpeakWeekday As Integer
        Dim NotpeakWeekend As Integer
        Dim Startpeak As Date = StartDate.Year.ToString + "/5/15"
        Dim Endpeak As Date = StartDate.Year.ToString + "/8/15"
        Dim Datecounter As String = ""


        PeakWeekday = 0
        PeakWeekend = 0
        NotpeakWeekday = 0
        NotpeakWeekend = 0

        '15/5~15/8
        count = Int((EndDate.Date - StartDate.Date).Days.ToString)
        For i As Integer = 0 To count - 1
            If StartDate.Date >= Startpeak And StartDate <= Endpeak Then

                If isWeekend(StartDate) Then
                    PeakWeekend += 1
                    ' MsgBox(Startpeak)
                Else
                    PeakWeekday += 1



                End If
            Else
                If isWeekend(StartDate) Then
                    NotpeakWeekend += 1

                Else
                    NotpeakWeekday += 1



                End If
            End If
            StartDate = StartDate.AddDays(1)


        Next

        Datecounter &= PeakWeekday & "," & PeakWeekend & "," & NotpeakWeekday & "," & NotpeakWeekend


        Return Datecounter


    End Function
    Public Shared Function isWeekend(ByVal CheckDate As Date) As Boolean
        If CheckDate.DayOfWeek = 6 Or CheckDate.DayOfWeek = 5 Then
            Return True
        Else
            Return False

        End If


    End Function
    
   
    Public Shared Function SelectAvailableRooms(ByVal Startdate As Date, ByVal Enddate As Date, ByVal Count As Boolean) As String
        Dim sqlstr As String = ""
        If (Count) Then
            sqlstr = " SELECT COUNT(RoomID) AS Total,[Room Type] FROM ("
        End If
        sqlstr &= " SELECT * FROM RoomTable WHERE [Room Number] NOT IN"
        sqlstr &= " (SELECT RoomNumber FROM RoomReservation "
        sqlstr &= "  WHERE RoomStart BETWEEN #" + Startdate.Date + "# AND #" + Enddate.Date + "#"
        sqlstr &= "  OR RoomEnd BETWEEN #" + Startdate.Date + "# AND #" + Enddate.Date + "#"
        sqlstr &= "  OR (RoomStart<=#" + Startdate.Date + "# AND RoomEnd>=#" + Enddate.Date + "#)"
        sqlstr &= "  OR (RoomStart>=#" + Startdate.Date + "# AND RoomEnd<=#" + Enddate.Date + "#)"
        sqlstr &= " )"
        If (Count) Then
            sqlstr &= "  )GROUP BY [Room Type]"
        End If

        Return sqlstr
    End Function
    Public Shared Function SelectRoomsRATE(ByVal RoomType As String, ByVal NotPweekend As String, ByVal NotPweekday As String, ByVal Pweekend As String, ByVal Pweekday As String) As String
        Dim SqlStr As String = ""
        SqlStr &= " SELECT total+total2 as FinalCost, A.Type FROM ("
        SqlStr &= " SELECT price1+price2 as total ,AA.RoomType as Type from "
        ' notpeakweekend
        SqlStr &= " (SELECT price*" & NotPweekend & " as price1,TypeID,RoomType from RoomType where RoomType='" & RoomType & "' AND Weekend='Y' AND PeakSeason='N') AS AA"
        SqlStr &= " LEFT JOIN"
        ' not peakweekday
        SqlStr &= " (SELECT price*" & NotPweekday & " as price2, TypeID,RoomType from RoomType where RoomType='" & RoomType & "' AND Weekend='N' AND PeakSeason='N')AS  BB"
        SqlStr &= " ON  AA.RoomType=BB.RoomType"
        SqlStr &= " ) AS A"
        SqlStr &= " LEFT JOIN"



        SqlStr &= " ("
        SqlStr &= " SELECT price1+price2 as total2 ,AA.RoomType as Type from "
        ' peakweekend
        SqlStr &= " (SELECT price*" & Pweekend & " as price1,TypeID,RoomType from RoomType where RoomType='" & RoomType & "' AND Weekend='Y' AND PeakSeason='Y') AS AA"
        SqlStr &= " LEFT JOIN"
        ' peakweekday
        SqlStr &= " (SELECT price*" & Pweekday & " as price2, TypeID,RoomType from RoomType where RoomType='" & RoomType & "' AND Weekend='N' AND PeakSeason='Y')AS  BB"
        SqlStr &= " ON  AA.RoomType=BB.RoomType"
        SqlStr &= " ) AS B"
        SqlStr &= " ON A.Type=B.Type"

        Return SqlStr

    End Function


 

  
    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged


        If (DataGridView2.CurrentRow.Selected) Then
         

            Dim RoomType As String = DataGridView2.SelectedRows(0).Cells("Room Type").Value
            Dim tmp As String() = DateCheck(DateTimePicker1.Value.Date, DateTimePicker2.Value.Date).Split(",")
            'tmp(0) is how many days of peakweekday 
            'tmp(1) is peakweekend 
            'tmp (2) is notpeakweekday
            'tmp(3) is notpeakweekend
            ' Label4.Text = tmp(0) & "_" & tmp(1) & "_" & tmp(2) & "_" & tmp(3)
            SQLstr = ""

            SQLstr = SelectRoomsRATE(RoomType, tmp(3), tmp(2), tmp(1), tmp(0))

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)

            SqlDa.Fill(DatSt, "RATE")
            Label60.Text = "$" & DatSt.Tables("RATE").Rows(0)(0).ToString

            Label5.Text = tmp(3) & " Day(s) in Weekend " & tmp(2) & " Day(s) in Weekday"
            Label6.Text = tmp(1) & " Day(s) in Weekend(Peak) " & tmp(0) & " Day(s) in Weekday(Peak)"
            DatSt.Tables.Remove("RATE")
            SQLstr = ""
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If (DataGridView1.DataSource Is Nothing) Then
            MsgBox("Plz search available rooms first!")
            Return
        End If
        If (DateTimePicker1.Value > DateTimePicker2.Value) Then


            MsgBox("CheckInDate must greater than CheckOutDate")
            Return

        End If
        Dim tmp As Integer = MsgBox("Customer data exist?", MsgBoxStyle.YesNoCancel)


        If tmp = MsgBoxResult.Yes Then
            TabControl1.SelectedIndex = 1




        ElseIf tmp = MsgBoxResult.No Then
            TabControl1.SelectedIndex = 4
        Else

        End If
        ComboBox5.Items.Clear()

        If DataGridView1.SelectedRows.Count > 1 Then


            For i As Integer = 0 To DataGridView1.RowCount - 2
                If DataGridView1.Rows(i).Selected = True Then
                    ComboBox5.Items.Add(DataGridView1.Rows(i).Cells("Room Number").Value)
                End If

            Next
        End If

        TextBox2.Text = DataGridView1.SelectedRows(0).Cells("Room Number").Value
        TextBox3.Text = DateTimePicker1.Value.Date
        TextBox4.Text = DateTimePicker2.Value.Date
        TextBox5.Text = Label4.Text.Remove(0, 1)


        If (TextBox32.Text <> "") Then
            TextBox21.Text = TextBox32.Text
        End If
        If (DataGridView1.SelectedRows.Count > 1) Then
            TextBox5.Text = Label61.Text
            TextBox2.Text = ""
            For i As Integer = 0 To ComboBox5.Items.Count - 1

                TextBox2.Text &= "," & ComboBox5.Items(i)
            Next
        End If

    End Sub

  

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

     

        Select Case ComboBox1.SelectedIndex
            Case 0
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p1.png"

            Case 1
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p1.png"
            Case 2
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p2.png"
            Case 3
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p2.png"
            Case 4
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p3.png"
            Case 5
                PictureBox1.ImageLocation = Directory.GetCurrentDirectory() & "\image\p3.png"

        End Select
        
        
        '   PictureBox1.ImageLocation = "C:\Users\Administrator\Documents\Visual Studio 2010\Projects\WindowsApplication2\WindowsApplication2\image\image.jpeg"
        ComboBox2.Items.Clear()
        Dim GuidedHikeNormal As String() = {6, 9, 12, 16, 18}

        Dim TimeCase As String() = {5, 8, 11, 14, 17}




        If ComboBox1.SelectedValue.ToString <> "System.Data.DataRowView" Then
            SQLstr = ""
            SQLstr = "Select Cost from RanchService Where [Service ID]=" & ComboBox1.SelectedValue.ToString

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "RanchCost")

            SQLstr = ""

            Cost.Text = DatSt.Tables("RanchCost").Rows(0)(0).ToString

            DatSt.Tables.Remove("RanchCost")
        End If


        Select Case ComboBox1.SelectedIndex
            Case 0
                For i As Integer = 0 To GuidedHikeNormal.Length - 1
                    ComboBox2.Items.Add(GuidedHikeNormal(i) & ":00")
                Next
            Case 1, 3
                For i As Integer = 0 To TimeCase.Length - 1
                    ComboBox2.Items.Add(TimeCase(i) & ":00")
                Next
            Case 2

                For i As Integer = 0 To TimeCase.Length - 1
                    ComboBox2.Items.Add(TimeCase(i) & ":30")
                Next
            Case 4
                For i As Integer = 8 To 23
                    ComboBox2.Items.Add(i & ":00")
                    ComboBox2.Items.Add(i & ":15")
                    ComboBox2.Items.Add(i & ":30")
                    ComboBox2.Items.Add(i & ":45")
                Next
            Case 5
                For i As Integer = 8 To 23
                    ComboBox2.Items.Add(i & ":00")
                    ComboBox2.Items.Add(i & ":30")

                Next
            Case 6, 7
                For i As Integer = 0 To 23
                    ComboBox2.Items.Add(i & ":00")
                    ComboBox2.Items.Add(i & ":30")

                Next


        End Select

        ComboBox2.SelectedIndex = 0

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
       
        ' Dim x As String = "#" & DateTimePicker3.Value.Date & " " & ComboBox2.SelectedItem.ToString & "#"
        

        Dim Startdt As DateTime
        Dim Enddt As DateTime
        Dim pickSrart As DateTime

        Dim pickEnd As DateTime
        Dim permission As Boolean = False
        Dim serviceID As String = ComboBox1.SelectedValue
        Dim reservationID As String
        Dim guestID As String = ""
        Dim ReservationStatus As String = "IN"
        
        Try
            SQLstr = ""
            SQLstr &= " SELECT ServiceMins FROM RanchService WHERE [Service ID]=" & ComboBox1.SelectedValue

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "MinsSearch")

            pickSrart = "#" & DateTimePicker3.Value.Date & " " & ComboBox2.SelectedItem.ToString & "#"
            pickEnd = "#" & DateTimePicker3.Value.Date & " " & ComboBox2.SelectedItem.ToString & "#"
            pickEnd = pickEnd.AddMinutes(DatSt.Tables("MinsSearch").Rows(0)(0))


            DatSt.Tables.Remove("MinsSearch")


            SQLstr = " "
            SQLstr = "SELECT Name, Gname,GuestID,R.ReservationID FROM Customer as C, RoomReservation as R, Guest as G"
            SQLstr &= " WHERE C.CustomerID=R.CustomerID"
            SQLstr &= " AND G.ReservationID=R.ReservationID"
            SQLstr &= " AND R.Status='IN' "
            SQLstr &= " AND R.CheckOut='N'"
            SQLstr &= " AND R.RoomNumber='" & TextBox23.Text & "' "
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "NameSearch")
            If DatSt.Tables("NameSearch").Rows.Count = 0 Then
                MsgBox("Wrong Room Number!")
                Return
            End If
            For i As Integer = 0 To DatSt.Tables("NameSearch").Rows.Count - 1

                If (TextBox25.Text.ToLower = DatSt.Tables("NameSearch").Rows(i)(0).ToString.ToLower Or TextBox25.Text.ToLower = DatSt.Tables("NameSearch").Rows(i)(1).ToString.ToLower) Then

                    permission = True


                End If



            Next
            guestID = DatSt.Tables("NameSearch").Rows(0)(2).ToString



            If Not permission Then
                MsgBox("This person is not allowed to make the ranch service")
                Return


            End If

            SQLstr = " "
            SQLstr = " SELECT AppointmentTime,ServiceMins,R.ReservationID FROM  ServiceReservation as S,RoomReservation as R ,RanchService as RS"
            SQLstr &= "  where R.ReservationID=S.ReservationID  AND R.Status='IN' AND CheckOut='N' AND S.Status='IN' AND S.ServiceID=RS.[Service ID]"
            SQLstr &= " AND R.RoomNumber='" & TextBox23.Text & "'"

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "RanchSearch")
            For i As Integer = 0 To DatSt.Tables("RanchSearch").Rows.Count - 1
                Startdt = DatSt.Tables("RanchSearch").Rows(i)(0)
                Enddt = "#" & DatSt.Tables("RanchSearch").Rows(i)(0) & "#"
                Enddt = Enddt.AddMinutes(DatSt.Tables("RanchSearch").Rows(i)(1))

                'MsgBox(Startdt)
                'MsgBox(Enddt)
                'MsgBox(pickSrart)
                'MsgBox(pickEnd)
                'MsgBox(pickSrart > Startdt)
                'MsgBox(Enddt > pickEnd)

                If (Startdt >= pickSrart And pickEnd >= Enddt) Then
                    MsgBox("The customer has already reserved the service during this time")
                    Return

                ElseIf (pickSrart > Startdt And Enddt > pickEnd) Then
                    MsgBox("The customer has already reserved the service during this time")
                    Return
                ElseIf (Startdt < pickEnd And Enddt > pickEnd) Then
                    MsgBox("The customer has already reserved the service during this time")
                    Return
                ElseIf (pickSrart > Startdt And pickSrart < Enddt) Then

                    MsgBox("The customer has already reserved the service during this time")
                    Return

                End If

            Next
            Select Case ComboBox1.SelectedIndex
                Case 0, 1
                    If TextBox24.Text > 8 Then
                        MsgBox("Capacity for this Ranch service is 8, plz change the number of the guests")
                        Return

                    End If
                Case 2, 3
                    If TextBox24.Text > 2 Then
                        MsgBox("Capacity for this Ranch service is 2, plz change the number of the guests")
                        Return

                    End If

                Case 4, 5
                    If TextBox24.Text > 12 Then
                        MsgBox("Capacity for this Ranch service is 12, plz change the number of the guests")
                        Return

                    End If
                Case 6, 7
                    If TextBox24.Text > 6 Then
                        MsgBox("Capacity for this Ranch service is 6, plz change the number of the guests")
                        Return

                    End If
            End Select

            SQLstr = ""
            SQLstr &= " SELECT  Capacity-SUM(Numguests) as C FROM ServiceReservation as S, RanchService as R"
            SQLstr &= " where  AppointmentTime=#" & pickSrart.ToString("yyyy/MM/dd HH:mm:ss") & "#"
            SQLstr &= " AND R.[Service ID]=S.ServiceID"
            SQLstr &= " AND S.Status='IN'"
            SQLstr &= " AND  S.ServiceID=" & serviceID
            SQLstr &= " group by S.ServiceID,Capacity"



            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "CapacitySearch")
            If (DatSt.Tables("CapacitySearch").Rows.Count > 0) Then
                'MsgBox(DatSt.Tables("CapacitySearch").Rows(0)(0))
                'MsgBox(pickSrart.ToString("yyyy/MM/dd HH:mm:ss"))
                If (DatSt.Tables("CapacitySearch").Rows(0)(0) < Int(TextBox24.Text)) Then
                    Dim tmp As Integer = MsgBox("The capacity at this time is full. Would you like to put in  the wait list?", MsgBoxStyle.YesNo)

                    If tmp = MsgBoxResult.Yes Then
                        ReservationStatus = "Wait"
                    ElseIf tmp = MsgBoxResult.No Then
                        DatSt.Tables.Remove("NameSearch")
                        DatSt.Tables.Remove("RanchSearch")
                        DatSt.Tables.Remove("CapacitySearch")
                        Return
                    Else

                    End If
                End If

            End If





            reservationID = DatSt.Tables("NameSearch").Rows(0)(3).ToString

            DatSt.Tables.Remove("NameSearch")
            DatSt.Tables.Remove("RanchSearch")
            DatSt.Tables.Remove("CapacitySearch")

            SQLstr = ""

            SQLstr &= " Insert INTO ServiceReservation (ServiceID,ReservationID,AppointmentTime,NumGuests,Status,[Timestamp],Cost,PaymentStatus,GuestID)"
            SQLstr &= " Values(" & serviceID & "," & reservationID & ",#" & pickSrart.ToString("yyyy/MM/dd HH:mm:ss") & "#,'" & TextBox24.Text & "','" & ReservationStatus & "',#" & System.DateTime.Now().ToString("yyyy/MM/dd") & "#,'" & Cost.Text & "','Not Paid'," & guestID & ")"
            SqlCn.Open()
            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.InsertCommand = sqlcommand
            SqlDa.InsertCommand.ExecuteNonQuery()
            SqlCn.Close()

            Dim msg As String = "Reservation Type: " & ComboBox1.Text & "" & vbCrLf & " StartTime:" & pickSrart & ", EndTime:" & pickEnd & vbCrLf & "Cost:" & Cost.Text & " Status:" & ReservationStatus


            MsgBox(msg)
        Catch ex As Exception
            MsgBox("PlZ type all information!" & vbCrLf & ex.Message)
        End Try
    End Sub
   
  
  
   
 
   
   


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Percent As Double = 1
        Dim ReservationID As String = ""
       
        If DataGridView3.DataSource Is Nothing Then
            MsgBox("Plz Check record!")
            Return

        End If
        Select Case ComboBox3.SelectedIndex
            Case 0
                Percent = 1
            Case 1
                Percent = 0.75
            Case 2
                Percent = 0

        End Select
        SQLstr = ""
        SQLstr = "Select CustomerID, Cost from RoomReservation Where CustomerID=" & TextBox18.Text & "AND CheckIn='N' AND CheckOut='N' AND Status='IN'"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "CheckIn")
        SQLstr = ""
        If DataGridView3.ColumnCount = 0 Then
            MsgBox("Please check the Customer record")
            Return

        End If
        ReservationID = DataGridView3.SelectedRows(0).Cells("ReservationID").Value
        Dim tmp As Integer = MsgBox("Are you sure you want to cancel this revervation?", MsgBoxStyle.YesNo)

        If tmp = MsgBoxResult.No Then

            Return
        End If


        If DatSt.Tables("CheckIn").Rows.Count <> 0 Then

            SQLstr = ""
            SQLstr = "INSERT  into Cancellation (ReservationID,[Refund Type],[Refund Description],[Refund Percent],[TimeStamp])"
            SQLstr &= " SELECT ReservationID,'" & ComboBox3.SelectedItem.ToString & "','" & TextBox22.Text & "','" & ComboBox4.SelectedItem.ToString & "','" & System.DateTime.Now & "' FROM RoomReservation WHERE CustomerID=" & TextBox18.Text & " AND Status='IN'"
            SQLstr &= " AND  ReservationID=" & ReservationID

            SqlCn.Open()
            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.InsertCommand = sqlcommand

            SqlDa.InsertCommand.ExecuteNonQuery()
            SQLstr = ""
            SQLstr = " UPDATE RoomReservation"
            SQLstr &= " SET Status='Cancelled'"
            SQLstr &= " WHERE ReservationID=(SELECT ReservationID FROM RoomReservation WHERE CustomerID=" & TextBox18.Text & " AND Status='IN' AND ReservationID=" & ReservationID & " )"


            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.UpdateCommand = sqlcommand

            SqlDa.UpdateCommand.ExecuteNonQuery()
            SqlCn.Close()

            MsgBox("Reservation cancelled. The cost: " & DatSt.Tables("CheckIn").Rows(0)(1) & "" & vbLf & "Refund amount:" & DatSt.Tables("CheckIn").Rows(0)(1) * Percent)
            DataGridView3.DataSource = Nothing
            TextBox18.Text = Nothing
            TextBox22.Text = Nothing

        Else
            MsgBox("Customer data does not exist")
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Try
            SQLstr = ""
            SQLstr = "INSERT  into Customer (Name,Age,Address,City,State,Zip,Phone,Email)"
            SQLstr &= " values ('" & TextBox10.Text & "','" & TextBox11.Text & "','" & TextBox13.Text & "','" & TextBox12.Text
            SQLstr &= "','" & TextBox14.Text & "','" & TextBox15.Text & "','" & TextBox16.Text & "','" & TextBox17.Text
            SQLstr &= "')"
            ' SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlCn.Open()



            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.InsertCommand = sqlcommand

            SqlDa.InsertCommand.ExecuteNonQuery()
            SQLstr = ""
            SQLstr = "Select CustomerID from Customer Where Name='" & TextBox10.Text & "' AND Phone='" & TextBox16.Text & "'"
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "CustomerID")
            SQLstr = ""

            TextBox1.Text = DatSt.Tables("CustomerID").Rows(0)(0).ToString
            MsgBox("Insert successfully CustomerID is " & DatSt.Tables("CustomerID").Rows(0)(0).ToString)
            DatSt.Tables.Remove("CustomerID")
            TabControl1.SelectedIndex = 1
            SqlCn.Close()
        Catch ex As Exception
            MsgBox("PlZ type all information!" & vbCrLf & ex.Message)
        End Try
        TextBox10.Text = Nothing
        TextBox11.Text = Nothing
        TextBox12.Text = Nothing
        TextBox13.Text = Nothing
        TextBox14.Text = Nothing
        TextBox15.Text = Nothing
        TextBox16.Text = Nothing
        TextBox17.Text = Nothing


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If Not IsNumeric(TextBox1.Text) Then
                MsgBox("Numbers only for CustomerID!")
                Return



            End If
            SQLstr = ""
            SQLstr = "Select CustomerID from Customer Where CustomerID=" & TextBox1.Text
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "CustomerID")
            SQLstr = ""


            If DatSt.Tables("CustomerID").Rows.Count <> 0 Then
                If (ComboBox5.Items.Count > 1) Then
                    Dim rooms As String = "RoomNumber: " & vbCrLf
                    For i As Integer = 0 To ComboBox5.Items.Count - 1
                        SQLstr = ""
                        SQLstr = "INSERT  into RoomReservation (CustomerID,RoomNumber,CreditCardNum,CardExpDate,CardCRVCode,NumGuests,Status,[RoomStart],[RoomEnd],CheckIn,CheckOut,[TimeStamp],Cost,PaymentStatus)"
                        SQLstr &= " values ('" & TextBox1.Text & "','" & ComboBox5.Items(i) & "','" & TextBox6.Text & "','" & TextBox7.Text
                        SQLstr &= "','" & TextBox8.Text & "','" & TextBox21.Text & "','IN','" & TextBox3.Text & "','" & TextBox4.Text
                        SQLstr &= "','N','N','" & System.DateTime.Now.ToString & "','" & TextBox5.Text & "',' Not Paid"
                        SQLstr &= "')"

                        SqlCn.Open()
                        sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                        SqlDa.InsertCommand = sqlcommand

                        SqlDa.InsertCommand.ExecuteNonQuery()
                        rooms &= "," & ComboBox5.Items(i) & vbCrLf
                        SqlCn.Close()

                    Next
                    MsgBox("Successfully Reserved," & vbCrLf & rooms & vbCrLf & "From " & TextBox3.Text & " to " & TextBox4.Text)
                Else

                    SQLstr = ""
                    SQLstr = "INSERT  into RoomReservation (CustomerID,RoomNumber,CreditCardNum,CardExpDate,CardCRVCode,NumGuests,Status,[RoomStart],[RoomEnd],CheckIn,CheckOut,[TimeStamp],Cost,PaymentStatus)"
                    SQLstr &= " values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox6.Text & "','" & TextBox7.Text
                    SQLstr &= "','" & TextBox8.Text & "','" & TextBox21.Text & "','IN','" & TextBox3.Text & "','" & TextBox4.Text
                    SQLstr &= "','N','N','" & System.DateTime.Now.ToString & "','" & TextBox5.Text & "',' Not Paid"
                    SQLstr &= "')"

                    SqlCn.Open()
                    sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                    SqlDa.InsertCommand = sqlcommand

                    SqlDa.InsertCommand.ExecuteNonQuery()
                    MsgBox("Successfully Reserved" & vbCrLf & "From " & TextBox3.Text & " to " & TextBox4.Text)
                    DataGridView1.DataSource = Nothing

                End If
            Else
                MsgBox("This CustomerID doesn't exist!")
            End If
            SqlCn.Close()

            DatSt.Tables.Remove("CustomerID")
        Catch ex As Exception
            MsgBox("PlZ type all information!" & vbCrLf & ex.Message)
        End Try
        TextBox1.Text = Nothing
        TextBox2.Text = Nothing
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        TextBox7.Text = Nothing
        TextBox8.Text = Nothing
        TextBox21.Text = Nothing


    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If Not IsNumeric(TextBox9.Text) Then
            MsgBox("Numbers only!")
            TextBox9.Text = Nothing
            Return


        End If
        Try
            SQLstr = ""
            SQLstr = "Select CustomerID,RoomStart,RoomEnd,RoomNumber from RoomReservation Where CustomerID=" & TextBox9.Text & "AND CheckIn='N' AND CheckOut='N' AND Status='IN'"
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "CheckIn")
            SQLstr = ""
            Dim RoomNumber As String = ""

            'MsgBox(ListBox1.Items.Count)
            'MsgBox(ListBox1.Items(1).ToString)
            'Return

            If DatSt.Tables("CheckIn").Rows.Count <> 0 Then
                '  For i As Integer = 0 To DatSt.Tables("CheckIn").Rows.Count - 1
                ' MsgBox(DatSt.Tables("CheckIn").Rows(0).Item("RoomNumber"))
                For i As Integer = 0 To DatSt.Tables("CheckIn").Rows.Count - 1
                    RoomNumber &= DatSt.Tables("CheckIn").Rows(i).Item("RoomNumber") & ","

                Next


                SQLstr = ""
                SQLstr = " UPDATE RoomReservation"
                SQLstr &= " SET CheckIn='Y',NumGuests='" & TextBox19.Text & "'"
                SQLstr &= " WHERE CheckOut='N'"
                SQLstr &= " AND  Status='IN'"
                SQLstr &= " AND CustomerID=" & TextBox9.Text
                SQLstr &= " AND RoomStart=(SELECT Top 1 Roomstart From  RoomReservation where customerID=" & TextBox9.Text & " AND Status='IN')"
                '  SQLstr &= " AND ReservationID=" & DatSt.Tables("CheckIn").Rows(0)(i)

                SqlCn.Open()
                sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                SqlDa.UpdateCommand = sqlcommand
                SqlDa.UpdateCommand.ExecuteNonQuery()
                SqlCn.Close()

                '    Next
                SqlCn.Open()
                If ListBox1.Items.Count > 0 Then


                    For i As Integer = 0 To ListBox1.Items.Count - 1
                        SQLstr = ""
                        SQLstr &= " INSERT INTO Guest(ReservationID,Gname)"
                        SQLstr &= " SELECT ReservationID,'" & ListBox1.Items(i).ToString & "' FROM RoomReservation WHERE CustomerID=" & TextBox9.Text
                        SQLstr &= " AND Status='IN' AND CheckOut='N'"
                        sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                        SqlDa.InsertCommand = sqlcommand
                        SqlDa.InsertCommand.ExecuteNonQuery()
                    Next
                Else
                    SQLstr = ""
                    SQLstr &= " INSERT INTO Guest(ReservationID,Gname)"
                    SQLstr &= " SELECT ReservationID,'None' FROM RoomReservation WHERE CustomerID=" & TextBox9.Text
                    SQLstr &= " AND Status='IN' AND CheckOut='N'"
                    sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                    SqlDa.InsertCommand = sqlcommand
                    SqlDa.InsertCommand.ExecuteNonQuery()
                End If
               
                MsgBox("Successfully Check In" & vbCrLf & " from " & DatSt.Tables("CheckIn").Rows(0)(1) & " to " & DatSt.Tables("CheckIn").Rows(0)(2) & vbCrLf & "RoomNumber:" & RoomNumber)

            Else
                MsgBox("no reservations found for the customer id")
            End If
            SqlCn.Close()
            TextBox9.Text = Nothing
            TextBox19.Text = Nothing
            ListBox1.Items.Clear()

            DatSt.Tables.Remove("CheckIn")
        Catch ex As Exception
            MsgBox("PlZ type all information!" & vbCrLf & ex.Message)
        End Try
    End Sub

   

    

   

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        SQLstr = ""

        SQLstr = " SELECT C.CustomerID,RoomNumber FROM Customer as C,RoomReservation as R WHERE Phone='" & TextBox30.Text & "'"
        SQLstr &= " AND R.CustomerID=C.CustomerID "
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "IDSearch")
        If DatSt.Tables("IDSearch").Rows.Count > 0 Then
            Label51.Text = "The Customer ID is :" & DatSt.Tables("IDSearch").Rows(0)(0)
            Label59.Text = " The Room Number is: " & DatSt.Tables("IDSearch").Rows(0)(1)
        Else
            MsgBox("The data dose not exist")
            Return
        End If
        SQLstr = ""
        SQLstr = " SELECT * FROM RoomReservation as R WHERE CustomerID=" & DatSt.Tables("IDSearch").Rows(0)(0) & ""
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "Search")
        DataGridView7.DataSource = DatSt.Tables("Search")
        SQLstr = ""
        SQLstr = " SELECT RS.[Service Type],S.Appointmenttime,S.Status,S.NumGuests,S.Cost FROM ServiceReservation as S,RoomReservation as R,RanchService as RS"
        SQLstr &= " WHERE S.ReservationID=R.ReservationID"
        SQLstr &= " AND S.ServiceID=RS.[Service ID]"
        SQLstr &= "  AND R.CustomerID=" & DatSt.Tables("IDSearch").Rows(0)(0)

        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "RanchSearch")
        DataGridView8.DataSource = DatSt.Tables("RanchSearch")

        DatSt.Tables.Remove("Search")
        DatSt.Tables.Remove("RanchSearch")
        DatSt.Tables.Remove("IDSearch")
    End Sub

   
   

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        If TextBox32.Text = "" Then
            MsgBox("Please type number of ppl")
            Return

        End If
        Dim group As New Form2
        Dim SingleR As String = "0"

        Dim DoubleR As String = "0"
        Dim Quadruple As String = "0"


        For i As Integer = 0 To DataGridView2.RowCount - 2
            If DataGridView2.Rows(i).Cells(1).Value.ToString = "Single" Then
                SingleR = DataGridView2.Rows(i).Cells(0).Value.ToString

            End If
            If DataGridView2.Rows(i).Cells(1).Value.ToString = "Double" Then
                DoubleR = DataGridView2.Rows(i).Cells(0).Value.ToString

            End If
            If DataGridView2.Rows(i).Cells(1).Value.ToString = "Quadruple" Then
                Quadruple = DataGridView2.Rows(i).Cells(0).Value.ToString

            End If
        Next
   

        group.Label5.Text = SingleR
        group.Label6.Text = DoubleR
        group.Label7.Text = Quadruple

        If Int(TextBox32.Text) Mod 2 <> 0 Then
            MsgBox("Please input a even number")
            Return

        End If

        Try


            group.Label1.Text += TextBox32.Text & " people ?"
            group.Label12.Text = TextBox32.Text
            group.Show()

            group = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = 0 Then
            ComboBox4.SelectedIndex = 0
        ElseIf (ComboBox3.SelectedIndex = 1) Then
            ComboBox4.SelectedIndex = 1
        ElseIf (ComboBox3.SelectedIndex = 2) Then
            ComboBox4.SelectedIndex = 2
        End If
    End Sub

 
    
   
    Private Sub TabPage1_Click(sender As System.Object, e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs) Handles Button12.Click
        If Not IsNumeric(TextBox18.Text) Then
            MsgBox("Numbers only!")
            TextBox18.Text = Nothing
            Return


        End If

        If TextBox18.Text = "" Then
            MsgBox("Please type CustomerID!")
            Return
        End If

        SQLstr = ""
        SQLstr = "Select ReservationID,RoomNumber,Status,RoomStart,RoomEnd from RoomReservation Where CustomerID=" & TextBox18.Text & "AND CheckIn='N' AND CheckOut='N' AND Status='IN'"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "CheckRecord")
        SQLstr = ""



        If DatSt.Tables("CheckRecord").Rows.Count <> 0 Then
            DataGridView3.DataSource = DatSt.Tables("CheckRecord")


        Else

            MsgBox("Customer data does not exist")
        End If
        DatSt.Tables.Remove("CheckRecord")

    End Sub
    

    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        If TextBox33.Text = "" Then
            MsgBox("Please type roomnumber")
            Return

        End If
        SQLstr = ""
        SQLstr &= " Select RanchReservationID,[Service Type],AppointmentTime,S.NumGuests,S.Status,S.Cost,S.PaymentStatus,S.ServiceID from ServiceReservation as S, RoomReservation as R,RanchService as RS Where R.RoomNumber='" & TextBox33.Text & "' AND R.CheckIn='Y' AND R.CheckOut='N' AND R.Status='IN'"
        SQLstr &= " AND R.ReservationID=S.ReservationID AND S.ServiceID=RS.[Service ID] AND S.Status<>'Cancel'"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "CancelRanch")
        SQLstr = ""
        DataGridView4.DataSource = DatSt.Tables("CancelRanch")
        If DatSt.Tables("CancelRanch").Rows.Count = 0 Then
            MsgBox("No such record!")
            DataGridView4.DataSource = Nothing
            Return
        End If
        DatSt.Tables.Remove("CancelRanch")
    End Sub

    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs) Handles Button14.Click
        Dim percent As Double = 0
        Dim cost As Integer = 0
        Dim appDate As DateTime
        Dim Capacity As Integer = 0
        If DataGridView4.DataSource Is Nothing Then
            MsgBox("Plz search record first!")
            Return

        End If
        Select Case ComboBox7.SelectedIndex
            Case 0
                percent = 1
            Case 1
                percent = 0.75
            Case 2
                percent = 0

        End Select
      
   


        For i As Integer = 0 To DataGridView4.SelectedRows.Count - 1
            SqlCn.Open()
            appDate = DataGridView4.SelectedRows(i).Cells("AppointmentTime").Value.ToString
            

            cost = cost + Int(DataGridView4.SelectedRows(i).Cells("Cost").Value)
            SQLstr = ""
            SQLstr &= " INSERT INTO RanchServiceCancel(RanchReservationID,[Refund Type],[Refund Description],[Refund Percent],[TimeStamp])"
            SQLstr &= " Values (" & DataGridView4.SelectedRows(i).Cells("RanchReservationID").Value & ",'" & ComboBox6.SelectedItem.ToString & "','" & TextBox35.Text & "'," & percent * 100 & ",#" & System.DateTime.Now.ToString("yyyy/MM/dd") & "#)"

            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.InsertCommand = sqlcommand
            SqlDa.InsertCommand.ExecuteNonQuery()

            SQLstr = ""
            SQLstr = " UPDATE ServiceReservation"
            SQLstr &= " SET Status='Cancel'"
            SQLstr &= " WHERE Status='IN'"
            SQLstr &= " AND RanchReservationID=" & DataGridView4.SelectedRows(i).Cells("RanchReservationID").Value
            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.UpdateCommand = sqlcommand
            SqlDa.UpdateCommand.ExecuteNonQuery()
            SqlCn.Close()

            SQLstr = ""
            SQLstr &= " SELECT  Capacity FROM RanchService "
            SQLstr &= " WHERE  [Service ID]=" & DataGridView4.SelectedRows(i).Cells("ServiceID").Value
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "Capacity")
            Capacity = DatSt.Tables("Capacity").Rows(0)(0)
            DatSt.Tables.Remove("Capacity")

            SQLstr = ""
            SQLstr &= " SELECT  Capacity-SUM(Numguests) as C FROM ServiceReservation as S, RanchService as R"
            SQLstr &= " where   AppointmentTime=#" & appDate.ToString("yyyy/MM/dd hh:mm:ss") & "#"
            SQLstr &= " AND R.[Service ID]=S.ServiceID"
            SQLstr &= " AND S.Status='IN'"
            SQLstr &= " AND  S.ServiceID=" & DataGridView4.SelectedRows(i).Cells("ServiceID").Value
            SQLstr &= " group by S.ServiceID,Capacity"
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SQLstr = ""
            SqlDa.Fill(DatSt, "SearchList")
            If DatSt.Tables("SearchList").Rows.Count > 0 Then
                Capacity = DatSt.Tables("SearchList").Rows(0)(0)
            End If
           
            SQLstr &= " Select S.NumGuests,RoomNumber,S.Appointmenttime,RanchReservationID from ServiceReservation as S ,RoomReservation as R Where ServiceID=" & DataGridView4.SelectedRows(i).Cells("ServiceID").Value & " AND AppointmentTime=#" & appDate.ToString("yyyy/MM/dd hh:mm:ss") & "#"
            SQLstr &= " AND S.Status='Wait' "
            SQLstr &= " AND R.ReservationID=S.ReservationID"
            SQLstr &= " ORDER BY S.TimeStamp"
            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "WaitList")
            SQLstr = ""
            'MsgBox(appDate.ToString("yyyy/MM/dd hh:mm:ss"))
            'MsgBox(Capacity)
            'MsgBox(DatSt.Tables("WaitList").Rows(0)(0))
            If DatSt.Tables("WaitList").Rows.Count > 0 Then
                If Capacity >= Int(DatSt.Tables("WaitList").Rows(0)(0)) Then
                    SqlCn.Open()
                    SQLstr = ""
                    SQLstr = " UPDATE ServiceReservation"
                    SQLstr &= " SET Status='IN'"
                    SQLstr &= " WHERE Status='Wait'"
                    SQLstr &= " AND RanchReservationID=" & DatSt.Tables("WaitList").Rows(0)(3).ToString
                    sqlcommand = New OleDbCommand(SQLstr, SqlCn)
                    SqlDa.UpdateCommand = sqlcommand
                    SqlDa.UpdateCommand.ExecuteNonQuery()
                    SqlCn.Close()
                    MsgBox("Update Status from WAIT to IN: RoomNumber: " & DatSt.Tables("WaitList").Rows(0)(1).ToString & " AppointmentTime: " & DatSt.Tables("WaitList").Rows(0)(2).ToString)

                End If
            End If
           
            DatSt.Tables.Remove("WaitList")
            DatSt.Tables.Remove("SearchList")

        Next

        MsgBox("Successfully canceled:" & vbCrLf & "Refund=" & cost * percent)
        DataGridView4.DataSource = Nothing
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        ComboBox7.SelectedIndex = ComboBox6.SelectedIndex

    End Sub

    
    Private Sub Button15_Click(sender As System.Object, e As System.EventArgs) Handles Button15.Click
        Dim Ranchcost As Double = 0
        Dim percent As Double = 0.0
        ComboBox8.SelectedIndex = 0
        Dim RoomNumberstr As String = " AND RoomNumber='" & TextBox31.Text & "'"
        If Not IsNumeric(TextBox29.Text) Then
            MsgBox("Numbers only for ID!")
            TextBox29.Text = Nothing
            Return


        End If

        SQLstr = ""
        SQLstr &= " Select * from RoomReservation "
        SQLstr &= " Where 1=1"
        If CheckBox1.Checked = False Then
            SQLstr &= RoomNumberstr
        End If
        SQLstr &= " AND CustomerID=" & TextBox29.Text
        SQLstr &= " AND CheckIn='Y'"
        SQLstr &= " AND CheckOut='N'"
        SQLstr &= " AND Status='IN'"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "RoomSummary")
        SQLstr = ""
        DataGridView5.DataSource = DatSt.Tables("RoomSummary")
        If DatSt.Tables("RoomSummary").Rows.Count = 0 Then
            MsgBox("No such record!Plz check ID&Room Number")
            DataGridView5.DataSource = Nothing
            Return

        End If
        Label68.Text = DataGridView5.SelectedRows(0).Cells("cost").Value
        DatSt.Tables.Remove("RoomSummary")

        SQLstr = ""
        SQLstr &= "SELECT * FROM "
        SQLstr &= " (Select S.NumGuests,RoomNumber,S.cost,S.Appointmenttime,[Service Type],S.Status,S.RanchReservationID from ServiceReservation as S ,RoomReservation as R ,RanchService as ST"

        SQLstr &= " Where S.ServiceID=ST.[Service ID] AND S.status<>'Wait'"

        SQLstr &= " AND R.ReservationID=S.ReservationID"
        If CheckBox1.Checked = False Then
            SQLstr &= RoomNumberstr
        End If
        SQLstr &= " AND R.CustomerID=" & TextBox29.Text
        SQLstr &= " ) as A LEFT JOIN  (Select RanchReservationID,[Refund Type],[Refund Percent] FROM RanchServiceCancel) as B "
        SQLstr &= " ON A.RanchReservationID=B.RanchReservationID"
        SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
        SqlDa.Fill(DatSt, "RanchSummary")
        SQLstr = ""
        DataGridView6.DataSource = DatSt.Tables("RanchSummary")
      
        For i As Integer = 0 To DatSt.Tables("RanchSummary").Rows.Count - 1
            'MsgBox(DatSt.Tables("RanchSummary").Rows(i).Item("cost"))
            'If DatSt.Tables("RanchSummary").Rows(i).Item("cost") = "" Then
            'Ranchcost += DatSt.Tables("RanchSummary").Rows(i).Item("cost")
            'End If
            If IsDBNull(DatSt.Tables("RanchSummary").Rows(i).Item("Refund Type")) Then
                Ranchcost += (DatSt.Tables("RanchSummary").Rows(i).Item("cost") * 1)
            Else
                percent = DatSt.Tables("RanchSummary").Rows(i).Item("Refund percent") / 100
                Ranchcost += (DatSt.Tables("RanchSummary").Rows(i).Item("cost") * percent)

            End If
        Next
        Label70.Text = Ranchcost

        DataGridView6.Columns("A.RanchReservationID").Visible = False
        DataGridView6.Columns("B.RanchReservationID").Visible = False

        DatSt.Tables.Remove("RanchSummary")
    End Sub

    
    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        Dim Checkout As New Form3

        Dim RoomNumber As String = "AND RoomNumber='" & TextBox31.Text & "'"
        Try
            If DataGridView5.DataSource Is Nothing Then
                MsgBox("PlZ check Summary first!")
                Return

            End If
            SqlCn.Open()
            SQLstr = ""
            SQLstr = " UPDATE RoomReservation"
            SQLstr &= " SET Checkout='Y',PaymentStatus='Paid'"
            SQLstr &= " WHERE 1=1"
            If CheckBox1.Checked = False Then
                SQLstr &= RoomNumber
            End If
            SQLstr &= " AND CustomerID=" & TextBox29.Text
            SQLstr &= " AND CheckIn='Y'"
            SQLstr &= " AND Status<>'Cancel'"
            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.UpdateCommand = sqlcommand
            SqlDa.UpdateCommand.ExecuteNonQuery()
            SQLstr = ""

            SQLstr = " UPDATE ServiceReservation"
            SQLstr &= " SET PaymentStatus='Paid'"
            SQLstr &= " WHERE ReservationID=" & DataGridView5.SelectedRows(0).Cells("ReservationID").Value

            sqlcommand = New OleDbCommand(SQLstr, SqlCn)
            SqlDa.UpdateCommand = sqlcommand
            SqlDa.UpdateCommand.ExecuteNonQuery()
            SQLstr = ""
            SqlCn.Close()
            MsgBox("Successfully checking out ! ")
            SQLstr = ""
            SQLstr &= " SELECT Name FROM Roomreservation as R,Customer as C"
            SQLstr &= " WHERE R.customerID=C.customerID"
            SQLstr &= " AND R.ReservationID=" & DataGridView5.SelectedRows(0).Cells("ReservationID").Value

            SqlDa = New OleDbDataAdapter(SQLstr, SqlCn)
            SqlDa.Fill(DatSt, "GetName")
            RoomNumber = ""
            Checkout.Label3.Text = DatSt.Tables("GetName").Rows(0)(0)
            If DataGridView5.RowCount > 2 Then

                For i As Integer = 1 To DataGridView5.RowCount - 2
                    RoomNumber &= "," & DataGridView5.Rows(i).Cells("RoomNumber").Value
                Next
            End If
            Checkout.Label6.Text = DataGridView5.SelectedRows(0).Cells("RoomNumber").Value & RoomNumber

            Checkout.Label7.Text = DataGridView5.SelectedRows(0).Cells("RoomStart").Value
            Checkout.Label9.Text = DataGridView5.SelectedRows(0).Cells("RoomEnd").Value
            Checkout.Label11.Text = Int(Label70.Text) + Int(Label68.Text)
            Checkout.Label11.Text &= "$"

            Checkout.Label14.Text = (Int(Label70.Text) + Int(Label68.Text)) * 0.1
            Checkout.Label14.Text &= "$"
            Checkout.Label13.Text = ComboBox8.SelectedItem.ToString



            Checkout.Show()
        Catch ex As Exception
            MsgBox("PlZ type all information!" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If TextBox31.Visible = True Then
            TextBox31.Visible = False

        Else
            TextBox31.Visible = True

        End If

    End Sub

    Private Sub Button16_Click(sender As System.Object, e As System.EventArgs) Handles Button16.Click
        ListBox1.Items.Add(TextBox20.Text)
        TextBox20.Text = Nothing

    End Sub

 

    Private Sub Button17_Click(sender As System.Object, e As System.EventArgs) Handles Button17.Click
        ListBox1.Items.Clear()



    End Sub
    

    

End Class
