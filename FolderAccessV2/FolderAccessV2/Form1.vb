Imports System.Data.SqlClient
Imports System.Security.AccessControl
Imports System.IO
Imports outlook = Microsoft.Office.Interop.Outlook
Public Class Form1

    Dim connection As New SqlConnection("Server= INNOWIN43169; Database = DEV_REPORTING; Integrated Security = true")
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For Each item In ListBox2.Items

            Call ExtractData(item)

        Next

        MsgBox("Access Granted successfully", vbInformation)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim i As Integer

        If ListBox2.Items.Count > 0 Then

            i = MsgBox("It looks like you have been editing something." & vbNewLine & vbNewLine & " Are you sure you want to exit?", vbYesNo + vbInformation, "Exit Application")

            If i = 6 Then
                Me.Close()
            Else
                Exit Sub
            End If
        Else
            i = MsgBox(" Are you sure you want to exit?", vbYesNo + vbInformation, "Exit Application")

            If i = 6 Then
                Me.Close()
            Else
                Exit Sub
            End If

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'ListBox1.Items.Add("hello")
        'ListBox1.Items.Add("hello1")
        'ListBox1.Items.Add("hello2")
        'ListBox1.Items.Add("hello3")
        'ListBox1.Items.Add("Dumy_Folder")

        Dim cmd As SqlCommand
        Try

            Dim sql As String = "Select * from T_RNA_FOLDER_DIRECTORY"
            connection.Open()
            cmd = New SqlCommand(sql, connection)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader
            Do While dr.Read
                ListBox1.Items.Add(dr.GetValue(0).ToString())
            Loop

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        connection.Close()
        'connection.Dispose()
        'connection = Nothing


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
        CheckBox1.Checked = False
        ListBox1.ClearSelected()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim selecteditems = (From i In ListBox1.SelectedItems).ToList

        If ListBox1.SelectedIndex = -1 Then
            MsgBox("No item is selected", vbExclamation)
            Exit Sub
        End If
        For Each selecteditem In selecteditems

            ListBox2.Items.Add(selecteditem)

        Next

        ListBox1.ClearSelected()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim selecteditems = (From i In ListBox2.SelectedItems).ToList

        If ListBox2.SelectedIndex = -1 Then
            MsgBox("No item is selected", vbExclamation)
            Exit Sub
        End If
        For Each selecteditem In selecteditems

            ListBox2.Items.Remove(selecteditem)

        Next

    End Sub

    Sub ExtractData(foldername As String)

        Dim searchresult As String

        Dim searchquery = "select * from T_RNA_FOLDER_DIRECTORY where [folder name] = " & "'" & foldername & "'"

        Dim command As New SqlCommand(searchquery, connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()

        adapter.Fill(table)

        searchresult = table.Rows(0)(1).ToString()

        Call grant_access(searchresult)

    End Sub

    Sub grant_access(FolderPath As String)

        Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
        Dim dSecurity As DirectorySecurity = FolderInfo.GetAccessControl()
        Dim UserAccount As String

        If TextBox1.Text <> "" Then
            UserAccount = TextBox1.Text
            If RadioButton1.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton2.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.ReadAndExecute, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton3.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Read, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton4.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Write, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton5.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            End If
            FolderInfo.SetAccessControl(dSecurity)
        End If

        If TextBox2.Text <> "" Then
            UserAccount = TextBox2.Text
            If RadioButton1.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton2.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.ReadAndExecute, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton3.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Read, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton4.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Write, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton5.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            End If
            FolderInfo.SetAccessControl(dSecurity)
        End If

        If TextBox3.Text <> "" Then
            UserAccount = TextBox3.Text
            If RadioButton1.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton2.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.ReadAndExecute, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton3.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Read, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton4.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Write, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton5.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            End If
            FolderInfo.SetAccessControl(dSecurity)
        End If

        If TextBox4.Text <> "" Then
            UserAccount = TextBox4.Text
            If RadioButton1.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton2.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.ReadAndExecute, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton3.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Read, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton4.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Write, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton5.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            End If
            FolderInfo.SetAccessControl(dSecurity)
        End If

        If TextBox5.Text <> "" Then
            UserAccount = TextBox5.Text
            If RadioButton1.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton2.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.ReadAndExecute, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton3.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Read, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton4.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Write, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            ElseIf RadioButton5.Checked = True Then
                dSecurity.AddAccessRule(New FileSystemAccessRule(UserAccount, FileSystemRights.Modify, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
            End If
            FolderInfo.SetAccessControl(dSecurity)
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Form2.Show()
        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim OutlookMessage As outlook.MailItem
        Dim AppOutlook As New outlook.Application

        Dim ind As Integer = 0


        Try
            Console.WriteLine("Starting 1")
            OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
            Dim Recipents As outlook.Recipients = OutlookMessage.Recipients
            Recipents.Add("ravjyot.singh.duggal@ericsson.com")
            OutlookMessage.Subject = "Testing"
            OutlookMessage.Body = "Testing Demo"
            OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
            Console.WriteLine("Before Send")
            'OutlookMessage.Send()
            OutlookMessage.Display()
            Console.WriteLine("Email Sent For " + "Ravjyot")

        Catch ex As Exception
            Console.WriteLine("Email Not Sent")
            'MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line 

        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
            ind = ind + 1
        End Try
        Console.Read()
    End Sub
End Class
