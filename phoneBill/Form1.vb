Public Class Form1
    Dim rateObj As New rateClass
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset    'connect to table rate
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'setting up the ADO objects
        conn.Provider = "Microsoft.Jet.OleDB.4.0"
        ' Setting up the jet DB driver (access) 
        conn.ConnectionString = "C:\ITD\Term 3\Visual Basic.Net\assignment\Midterm\phoneBillDB.mdb"
        conn.Open()
        rs.Open("select * from rateTable", conn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockPessimistic)
    End Sub
    'Adding data from form to DB
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or
           TextBox2.Text = "" Or
           TextBox3.Text = "" Then
            MessageBox.Show("Please Fill up all boxes")
            Exit Sub
        End If

        Dim Criteria As String

        Criteria = "countryID =" + TextBox1.Text
        rs.MoveFirst()
        'go to the beginning to start serach 
        rs.Find(Criteria)
        ' Either We find the record(s), which is the first record if there are more than one
        ' If record is found the file pointer stays at it
        'if not found, the file pointer has passed eof meaning eof = true
        If rs.EOF Then
            'not found
            rs.AddNew()
            Call SaveinTable()
            rs.Update()
            MessageBox.Show("Record Added succesfully")
            ' add items for combobox too
            ComboBox1.Items.Add(rs.Fields("countryName").Value)
            Call ClearForm()
            Exit Sub
        Else
            'found 
            Call showdata()
            MessageBox.Show("Duplicate Record, try another ID")
            Exit Sub
        End If

    End Sub
    'function for save any data in table
    Private Sub SaveinTable()
        rateObj.countryID = Convert.ToInt32(TextBox1.Text)
        rateObj.countryName = TextBox2.Text
        rateObj.rate = Convert.ToInt32(TextBox3.Text)
        rs.Fields("countryID").Value = rateObj.countryID
        rs.Fields("countryName").Value = rateObj.countryName
        rs.Fields("rate").Value = rateObj.rate
    End Sub
    ' database to obj
    Public Sub moveObj()
        rateObj.countryID = rs.Fields("countryID").Value
        rateObj.countryName = rs.Fields("countryName").Value
        rateObj.rate = rs.Fields("rate").Value
    End Sub
    'show data Obj in form
    Private Sub showdata()
        TextBox1.Text = rateObj.countryID
        TextBox2.Text = rateObj.countryName
        TextBox3.Text = rateObj.rate
    End Sub

    Private Sub ClearForm()
        TextBox1.Text = ""
        TextBox2.Clear()
        TextBox3.Text = ""
    End Sub
    'Modifying data from form to DB
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or
           TextBox2.Text = "" Or
           TextBox3.Text = "" Then
            MessageBox.Show("Please Fill up all boxes")
            Exit Sub
        End If
        Dim Criteria As String
        Criteria = "countrytID =" + TextBox1.Text

        rateObj.countryID = Convert.ToInt32(TextBox1.Text)
        rateObj.countryName = TextBox2.Text
        rateObj.rate = Convert.ToInt32(TextBox3.Text)

        'go to the beginning to start serach 
        rs.Find(Criteria)
        ' Either We find the record(s), which is the first record if there are more than one
        ' If record is found the file pointer stays at it
        'if not found, the file pointer has passed eof meaning eof = true
        If rs.EOF Then
            ' it is impossible, if you refrain from changing the ID 
            Call showdata()
            MessageBox.Show("Record with this ID does not exist")
            Exit Sub
        Else
            'found 
            Call SaveinTable()
            ' delete items from combobox anyways
            ComboBox1.Items.Remove(rs.Fields("countryName").Value)
            rs.Update()
            MessageBox.Show("Record Modified succesfully")
            ComboBox1.Items.Add(rateObj.countryName)
        End If
    End Sub
    ' searching data from DB
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Criteria As String
        Criteria = ""
        If (TextBox1.Text <> "") Then
            Criteria = Criteria + "countryID = " + TextBox1.Text
        End If
        If (TextBox2.Text <> "") Then
            If Criteria <> "" Then
                Criteria = Criteria + " AND countryName = '" + TextBox2.Text + "'"
            Else
                Criteria = Criteria + "countryName = '" + TextBox2.Text + "'"
            End If
        End If
        If (TextBox3.Text <> "") Then
            If Criteria <> "" Then
                Criteria = Criteria + " AND rate = '" + TextBox3.Text + "'"
            Else
                Criteria = Criteria + "rate = '" + TextBox3.Text + "'"
            End If
        End If
        'MessageBox.Show(Criteria)
        rs.MoveFirst()
        rs.Filter = Criteria
        If rs.EOF Then
            'not found
            MessageBox.Show("Recod with your specific criteria not found")
            Exit Sub
        Else
            Call showdata()
            rs.Filter = ""
        End If
    End Sub

    ' delete button
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("Please enter the ID you want to delete")
            Exit Sub
        End If

        Dim Criteria As String
        Criteria = "countryID = " + TextBox1.Text

        ' data that user input have already or not
        rs.Find(Criteria)
        If rs.EOF Then
            ' this account number is empty
            Call showdata()
            MessageBox.Show("Record didnt exist, please enter other one")
            Exit Sub
        Else
            ' found the record
            ' delete items from combobox 
            ComboBox1.Items.Remove(rs.Fields("countryName").Value)

            rs.Delete()
            MessageBox.Show("Record deleted successfully")
            'Call showFirst()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Please fill the form")
            Exit Sub
        End If

        rs.MoveFirst()

        Dim test As String
        test = "countryName = '" + ComboBox1.Text + "'"

        rs.Find(test)

        If rs.EOF Then
            ' this account number is empty
            rs.MoveFirst()
            Call moveObj()
            Call showdata()
            MessageBox.Show("Record didnt exist, please enter other one")
            Exit Sub
        Else
            ' found country
            Call moveObj()

            ' calculate 
            Dim cost As Integer
            cost = rateObj.rate * Convert.ToInt32(TextBox4.Text)

            TextBox5.Text = cost.ToString + "$"
            MessageBox.Show("Calculate Successfully")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' check the table is empty or not
        If rs.EOF = True And rs.BOF = True Then
            MessageBox.Show("Table is empty")
            Exit Sub
        End If
        rs.MovePrevious()
        ' if user put the previous button when the data is first message box show up
        If rs.BOF Then
            MessageBox.Show("Beginning of the table!")
            rs.MoveNext()
        End If
        Call moveObj()
        Call showdata()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' check the table is empty or not
        If rs.EOF = True And rs.BOF = True Then
            MessageBox.Show("Table is empty")
            Exit Sub
        End If
        rs.MoveNext()
        ' if user put the next button when the data is last message box show up
        If rs.EOF Then
            MessageBox.Show("End of the table!")
            rs.MovePrevious()
        End If
        Call moveObj()
        Call showdata()
    End Sub
End Class
