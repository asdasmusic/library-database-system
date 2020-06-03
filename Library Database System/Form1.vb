Imports System.IO

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListView1.Items.Count = 0 And TextBox1.Text <> String.Empty And IsNumeric(TextBox1.Text) = True And TextBox2.Text <> String.Empty And DateTimePicker1.Value.Date <= DateTimePicker2.Value Then
            Dim item1 As ListViewItem
            item1 = ListView1.Items.Add(TextBox1.Text)
            item1.SubItems.Add(TextBox2.Text)
            item1.SubItems.Add(TextBox3.Text)
            item1.SubItems.Add(TextBox4.Text)
            item1.SubItems.Add(TextBox5.Text)
            item1.SubItems.Add(DateTimePicker1.Text)
            item1.SubItems.Add(DateTimePicker2.Text)
            item1 = Nothing
        ElseIf TextBox1.Text <> String.Empty And IsNumeric(TextBox1.Text) = True And TextBox2.Text <> String.Empty And DateTimePicker1.Value.Date <= DateTimePicker2.Value.Date Then
            With ListView1
                Dim additem As ListViewItem
                Dim titleitem As String
                Dim nameitem As String
                Dim dateitem As String
                additem = .FindItemWithText(TextBox1.Text, True, 0, True)
                If Not additem Is Nothing Then
                    titleitem = additem.SubItems(1).Text
                    nameitem = additem.SubItems(4).Text
                    dateitem = additem.SubItems(5).Text
                    MessageBox.Show("This Book ID [" & TextBox1.Text & "] with Book Title """ & titleitem & """ has already been Issued to " & nameitem & " on " & dateitem & ".", "Cannot Add Entry", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    Dim item1 As ListViewItem
                    item1 = ListView1.Items.Add(TextBox1.Text)
                    item1.SubItems.Add(TextBox2.Text)
                    item1.SubItems.Add(TextBox3.Text)
                    item1.SubItems.Add(TextBox4.Text)
                    item1.SubItems.Add(TextBox5.Text)
                    item1.SubItems.Add(DateTimePicker1.Text)
                    item1.SubItems.Add(DateTimePicker2.Text)
                    item1 = Nothing
                End If
            End With
        End If
        If TextBox1.Text <> String.Empty And IsNumeric(TextBox1.Text) = False Then
            MessageBox.Show("Book ID should be a Number.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If TextBox1.Text = String.Empty And TextBox2.Text <> String.Empty Then
            MessageBox.Show("Book ID cannot be Blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If TextBox2.Text = String.Empty And TextBox1.Text <> String.Empty Then
            MessageBox.Show("Book Title cannot be Blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If TextBox1.Text = String.Empty And TextBox2.Text = String.Empty Then
            MessageBox.Show("Book ID and Book Title cannot be Blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If DateTimePicker1.Value.Date > DateTimePicker2.Value.Date Then
            MessageBox.Show("Return Date cannot be older than Issue Date.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Private Sub clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        DateTimePicker1.ResetText()
        DateTimePicker2.ResetText()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim iexit As DialogResult
        iexit = MessageBox.Show("Are you sure? Please Save your work before Closing.", "Close Application", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If iexit = DialogResult.Yes Then
            Application.Exit()

        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim item1 As ListViewItem
        If ListView1.Items.Count > 0 Then
            If ListView1.SelectedIndices.Count = 0 Then
                MessageBox.Show("Select a Database Entry to Remove.", "Remove Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                item1 = ListView1.SelectedItems(0)
                item1.Remove()
            End If

        ElseIf ListView1.Items.Count = 0 Then
            MessageBox.Show("Unable to Remove Database Entry as the Library Database is Empty.", "Remove Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If ListView1.Items.Count <> 0 Then
            Dim sfd As SaveFileDialog = New SaveFileDialog
            sfd.DefaultExt = "lds"
            sfd.FileName = "MyLibrary"
            sfd.InitialDirectory = Application.StartupPath
            sfd.Filter = "Library Database files|*.lds|All files|*.*"
            sfd.Title = "Save Library Database file"
            If sfd.ShowDialog() = DialogResult.OK Then
                Dim mywriter As New IO.StreamWriter(sfd.FileName)
                mywriter.WriteLine("--<:: LIBRARY DATABASE SYSTEM (LDS) v1.5 ---- SmolApps Team 2018 ::>--")
                mywriter.WriteLine("")
                mywriter.WriteLine(ColumnHeader1.Text & " | " & ColumnHeader2.Text & " | " & ColumnHeader3.Text & " | " & ColumnHeader4.Text & " | " & ColumnHeader5.Text & " | " & ColumnHeader6.Text & " | " & ColumnHeader7.Text)
                mywriter.WriteLine("____________________________________________________________________________")
                mywriter.WriteLine("")
                For Each item1 As ListViewItem In ListView1.Items()
                    mywriter.WriteLine(item1.Text & " | " & item1.SubItems(1).Text & " | " & item1.SubItems(2).Text & " | " & item1.SubItems(3).Text & " | " & item1.SubItems(4).Text & " | " & item1.SubItems(5).Text & " | " & item1.SubItems(6).Text)
                Next
                mywriter.Close()
                MessageBox.Show("Library Database has been successfully Saved.", "Save Finished", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Unable to export as the Library Database is Empty.", "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ofd As OpenFileDialog = New OpenFileDialog
        ofd.DefaultExt = "lds"
        ofd.FileName = "MyLibrary"
        ofd.InitialDirectory = Application.StartupPath
        ofd.Filter = "Library Database files|*.lds|All files|*.*"
        ofd.Title = "Open Library Database file"
        If ofd.ShowDialog() = DialogResult.OK Then
            ListView1.Items.Clear()
            Try
                Dim MyStream As New StreamReader(ofd.FileName)
                Dim strTemp() As String
                For i = 1 To 5
                    MyStream.ReadLine()
                Next
                Do While MyStream.Peek <> -1
                    Dim LVitem As New ListViewItem
                    strTemp = MyStream.ReadLine.Split(New String() {" | "}, StringSplitOptions.None)
                    LVitem.Text = strTemp(0).ToString
                    ListView1.Items.Add(LVitem)
                    LVitem.SubItems.Add(strTemp(1).ToString)
                    LVitem.SubItems.Add(strTemp(2).ToString)
                    LVitem.SubItems.Add(strTemp(3).ToString)
                    LVitem.SubItems.Add(strTemp(4).ToString)
                    LVitem.SubItems.Add(strTemp(5).ToString)
                    LVitem.SubItems.Add(strTemp(5).ToString)
                Loop
                MyStream.Close()
            Catch ex As Exception
                ListView1.Items.Clear()
                MessageBox.Show("Unable to read the Library Database file. It may be either corrupt or invalid.", "Open Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MessageBox.Show("LIBRARY DATABASE SYSTEM (LDS) v1.5" & vbNewLine & vbNewLine & "This mini app was made by the SmolApps Team." & vbNewLine & vbNewLine & "Credits: Anshuman, Ankush, Annu, Aritra, Anushri." & vbNewLine & vbNewLine & "Copyright © SmolApps 2018. All Rights Reserved.", "About", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class
