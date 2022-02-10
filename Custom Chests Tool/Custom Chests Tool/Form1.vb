Imports System.IO.Compression

Public Class Form1
    Dim appPath As String = ""
    Dim bnpPath As String = ""
    Dim wrkPath As String = ""
    Dim mapPath As String = ""
    Dim endPath As String = ""
    Dim endName As String = ""
    Dim clsWork As clsMap = Nothing



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' We start here. Setting paths and getting the form ready.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        appPath = String.Format("{0}\", Environment.CurrentDirectory)
        bnpPath = String.Format("{0}\", Environment.CurrentDirectory) & "Custom_Chests_Tool.dat"
        wrkPath = String.Format("{0}\", Environment.CurrentDirectory) & "your_changes\"
        mapPath = String.Format("{0}\", Environment.CurrentDirectory) & "your_changes\logs\map.yml"
        endPath = String.Format("{0}\", Environment.CurrentDirectory) & "Your_ArmorChestEmporium_v3.bnp"
        endName = "Your_ArmorChestEmporium_v3.bnp"
        grpCustom.Enabled = False
        grpRefill.Enabled = False
        btnSave.Enabled = False
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' NEW button. Will prepare the working folder and unzip the BNP file.
    ''' If previous working folder is found, it will delete it!
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        Cursor = Cursors.WaitCursor
        actNew(sender, e)
        Cursor = Cursors.Default
    End Sub
    Private Sub actNew(sender As Object, e As EventArgs)

        'Mandatory question.
        If MsgBox("Are you sure to discard your previous changes and start again from scracth?", vbExclamation + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then Exit Sub
        grpCustom.Enabled = False
        grpRefill.Enabled = False
        btnSave.Enabled = False
        lblError1.Visible = False
        lblError2.Visible = False
        lblError3.Visible = False
        lblError4.Visible = False
        lblError5.Visible = False
        lblError6.Visible = False
        lblError7.Visible = False
        lblError8.Visible = False
        lblError9.Visible = False
        lblError10.Visible = False
        lblError11.Visible = False
        lblError12.Visible = False
        lblError13.Visible = False
        lblError14.Visible = False
        lblError15.Visible = False
        lblErrorA.Visible = False
        lblErrorB.Visible = False
        lblErrorC.Visible = False
        lblErrorD.Visible = False
        txtCustom1.Text = ""
        txtCustom2.Text = ""
        txtCustom3.Text = ""
        txtCustom4.Text = ""
        txtCustom5.Text = ""
        txtCustom6.Text = ""
        txtCustom7.Text = ""
        txtCustom8.Text = ""
        txtCustom9.Text = ""
        txtCustom10.Text = ""
        txtCustom11.Text = ""
        txtCustom12.Text = ""
        txtCustom13.Text = ""
        txtCustom14.Text = ""
        txtCustom15.Text = ""
        txtRefill1.Text = ""
        txtRefill2.Text = ""
        txtRefill3.Text = ""
        txtRefill4.Text = ""

        'Delete any previous working folder.
        'It could be a bit tricky sometimes because Windows Explorer tends to lock files for a while.
        Dim redo As Boolean = False
        If IO.Directory.Exists(wrkPath) Then
            Try
                IO.Directory.Delete(wrkPath, True)
            Catch ex As Exception
                redo = True
            End Try
        End If
        If redo Then
            Threading.Thread.Sleep(2000)
            If IO.Directory.Exists(wrkPath) Then
                Try
                    IO.Directory.Delete(wrkPath, True)
                Catch ex As Exception
                    Threading.Thread.Sleep(2000)
                    If IO.Directory.Exists(wrkPath) Then
                        MsgBox("Unable to delete the working directory:" & vbCrLf & vbCrLf & wrkPath & vbCrLf & vbCrLf & ex.Message, vbCritical, "Error")
                        Exit Sub
                    End If
                End Try
            End If
        End If

        'Creating the new working folder.
        Threading.Thread.Sleep(2000)
        Try
            IO.Directory.CreateDirectory(wrkPath)
        Catch ex As Exception
            MsgBox("Unable to create working directory:" & vbCrLf & vbCrLf & wrkPath, vbCritical, "Error")
            Exit Sub
        End Try

        'Extract the source BNP in the working folder.
        'It could be a bit tricky because Windows Explorer seems to delay sometimes the folder creation.
        Dim lastError As String = ""
        For i As Integer = 1 To 3
            Threading.Thread.Sleep(2000)
            Try
                ZipFile.ExtractToDirectory(bnpPath, wrkPath)
                Exit For
            Catch ex As Exception
                lastError = ex.Message
            End Try
        Next
        Threading.Thread.Sleep(2000)
        If Not IO.File.Exists(mapPath) Then
            MsgBox("Unable to extract files to the working directory:" & vbCrLf & vbCrLf & lastError, vbCritical, "Error")
            Exit Sub
        End If

        'Now we call the loading part.
        btnLoad_Click(Nothing, Nothing)
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' LOAD button. Will read the map file and prepare the form for any user
    ''' interaction. Working folder and map file must exist. It can be called
    ''' with parameters set to Nothing: then no questions will be asked.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Cursor = Cursors.WaitCursor
        actLoad(sender, e)
        Cursor = Cursors.Default
    End Sub
    Private Sub actLoad(sender As Object, e As EventArgs)

        'Mandatory question, only if called via click event.
        If Not sender Is Nothing Then
            If MsgBox("Do you want to load your previous changes?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
        End If
        grpCustom.Enabled = False
        grpRefill.Enabled = False
        btnSave.Enabled = False
        lblError1.Visible = False
        lblError2.Visible = False
        lblError3.Visible = False
        lblError4.Visible = False
        lblError5.Visible = False
        lblError6.Visible = False
        lblError7.Visible = False
        lblError8.Visible = False
        lblError9.Visible = False
        lblError10.Visible = False
        lblError11.Visible = False
        lblError12.Visible = False
        lblError13.Visible = False
        lblError14.Visible = False
        lblError15.Visible = False
        lblErrorA.Visible = False
        lblErrorB.Visible = False
        lblErrorC.Visible = False
        lblErrorD.Visible = False
        txtCustom1.Text = ""
        txtCustom2.Text = ""
        txtCustom3.Text = ""
        txtCustom4.Text = ""
        txtCustom5.Text = ""
        txtCustom6.Text = ""
        txtCustom7.Text = ""
        txtCustom8.Text = ""
        txtCustom9.Text = ""
        txtCustom10.Text = ""
        txtCustom11.Text = ""
        txtCustom12.Text = ""
        txtCustom13.Text = ""
        txtCustom14.Text = ""
        txtCustom15.Text = ""
        txtRefill1.Text = ""
        txtRefill2.Text = ""
        txtRefill3.Text = ""
        txtRefill4.Text = ""

        'Check if the map file is ready.
        If Not IO.File.Exists(mapPath) Then
            MsgBox("Unable to find the map file in the working directory (if it's the first time you run this tool, please click the Start from scratch button):" & vbCrLf & vbCrLf & mapPath, vbCritical, "Error")
            Exit Sub
        End If

        'Load the map file.
        clsWork = New clsMap
        If Not clsWork.Load(mapPath) Then
            clsWork = Nothing
            Exit Sub
        End If

        'Simple verification...
        If clsWork.GetLine(292) = "ERROR" OrElse clsWork.GetLine(328) = "ERROR" OrElse clsWork.GetLine(574) = "ERROR" Then
            MsgBox("The map file is wrong or its content don't match with the expected values. Starting from scratch is suggested.", vbExclamation, "Error")
            clsWork = Nothing
            Exit Sub
        End If

        'Put in the form the contents of the working class.
        txtCustom1.Text = clsWork.GetLine(292)
        txtCustom2.Text = clsWork.GetLine(298)
        txtCustom3.Text = clsWork.GetLine(304)
        txtCustom4.Text = clsWork.GetLine(310)
        txtCustom5.Text = clsWork.GetLine(316)
        txtCustom6.Text = clsWork.GetLine(322)
        txtCustom7.Text = clsWork.GetLine(328)
        txtCustom8.Text = clsWork.GetLine(334)
        txtCustom9.Text = clsWork.GetLine(340)
        txtCustom10.Text = clsWork.GetLine(346)
        txtCustom11.Text = clsWork.GetLine(352)
        txtCustom12.Text = clsWork.GetLine(358)
        txtCustom13.Text = clsWork.GetLine(364)
        txtCustom14.Text = clsWork.GetLine(370)
        txtCustom15.Text = clsWork.GetLine(376)
        txtRefill1.Text = clsWork.GetLine(556)
        txtRefill2.Text = clsWork.GetLine(562)
        txtRefill3.Text = clsWork.GetLine(568)
        txtRefill4.Text = clsWork.GetLine(574)

        'Form ready for the user to start editing things.
        grpCustom.Enabled = True
        grpRefill.Enabled = True
        btnSave.Enabled = True
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' SAVE button. Will update the map file and zip the whole working folder.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Cursor = Cursors.WaitCursor
        actSave(sender, e)
        Cursor = Cursors.Default
    End Sub
    Private Sub actSave(sender As Object, e As EventArgs)
        Dim toDelete As Boolean = False

        'Check if the map file is ready.
        If Not IO.File.Exists(mapPath) Then
            MsgBox("Map file not found in the working directory! Did you erase it? Did you changed anything in the working folder? Starting from scratch is suggested." & vbCrLf & vbCrLf & mapPath, vbCritical, "Error")
            Exit Sub
        End If

        'Check if the target BNP already exist.
        If IO.File.Exists(endPath) Then
            If MsgBox("The mod file ''" & endName & "'' already exists. Overwrite it?" & vbCrLf & vbCrLf & endPath, vbExclamation + vbYesNo + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
            toDelete = True
        Else
            If MsgBox("Do you want to save your changes and generate the mod file ''" & endName & "''?", vbQuestion + vbYesNo + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
        End If

        'Update contents from the form in the working class.
        lblError1.Visible = Not clsWork.SetLine(292, txtCustom1.Text)
        lblError2.Visible = Not clsWork.SetLine(298, txtCustom2.Text)
        lblError3.Visible = Not clsWork.SetLine(304, txtCustom3.Text)
        lblError4.Visible = Not clsWork.SetLine(310, txtCustom4.Text)
        lblError5.Visible = Not clsWork.SetLine(316, txtCustom5.Text)
        lblError6.Visible = Not clsWork.SetLine(322, txtCustom6.Text)
        lblError7.Visible = Not clsWork.SetLine(328, txtCustom7.Text)
        lblError8.Visible = Not clsWork.SetLine(334, txtCustom8.Text)
        lblError9.Visible = Not clsWork.SetLine(340, txtCustom9.Text)
        lblError10.Visible = Not clsWork.SetLine(346, txtCustom10.Text)
        lblError11.Visible = Not clsWork.SetLine(352, txtCustom11.Text)
        lblError12.Visible = Not clsWork.SetLine(358, txtCustom12.Text)
        lblError13.Visible = Not clsWork.SetLine(364, txtCustom13.Text)
        lblError14.Visible = Not clsWork.SetLine(370, txtCustom14.Text)
        lblError15.Visible = Not clsWork.SetLine(376, txtCustom15.Text)
        lblErrorA.Visible = Not clsWork.SetLine(556, txtRefill1.Text)
        lblErrorB.Visible = Not clsWork.SetLine(562, txtRefill2.Text)
        lblErrorC.Visible = Not clsWork.SetLine(568, txtRefill3.Text)
        lblErrorD.Visible = Not clsWork.SetLine(574, txtRefill4.Text)
        If lblError1.Visible Or lblError2.Visible Or lblError3.Visible Or lblError4.Visible Or lblError5.Visible Or lblError6.Visible Or lblError7.Visible Or lblError8.Visible Or lblError9.Visible Or lblError10.Visible Or lblError11.Visible Or lblError12.Visible Or lblError13.Visible Or lblError14.Visible Or lblError15.Visible Or lblErrorA.Visible Or lblErrorB.Visible Or lblErrorC.Visible Or lblErrorD.Visible Then
            MsgBox("Some fields are not valid. Check the ones with red !! mark and try again.", vbExclamation, "Error")
            Exit Sub
        End If

        'If existing BNP deletion is needed, now it's time to do it.
        If toDelete Then
            Try
                IO.File.Delete(endPath)
            Catch ex As Exception
                MsgBox("Unable to delete the old mod file. Try to delete it by yourself and try again." & vbCrLf & vbCrLf & ex.Message, vbExclamation, "Error")
                Exit Sub
            End Try
            Threading.Thread.Sleep(2000)
        End If

        'Write the map file.
        If Not clsWork.Save(mapPath) Then
            clsWork = Nothing
            Exit Sub
        End If

        'ZIP the mod!
        Try
            ZipFile.CreateFromDirectory(wrkPath, endPath)
        Catch ex As Exception
            MsgBox("Unable while extracting files to the working directory:" & vbCrLf & vbCrLf & ex.Message, vbCritical, "Error")
            Exit Sub
        End Try

        'We're done!
        MsgBox("The mod file ''" & endName & "'' has been created! You can find it here and install/update on BCML:" & vbCrLf & vbCrLf & appPath, vbInformation, "Finished")
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Open the a page in the default browser.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub linkContents_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles linkContents.LinkClicked
        Try
            System.Diagnostics.Process.Start("https://github.com/MrCheeze/botw-tools/blob/master/botw_names.json")
        Catch ex As Exception
            MsgBox("Unable to open the page:" & ex.Message, vbExclamation, "Error")
        End Try
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Open the a page in the default browser.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub linkMod_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles linkMod.LinkClicked
        Try
            System.Diagnostics.Process.Start("https://gamebanana.com/mods/354876")
        Catch ex As Exception
            MsgBox("Unable to open the page:" & ex.Message, vbExclamation, "Error")
        End Try
    End Sub



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Dirty trick to select the contents when the user focuses it.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------
    Private Sub txtCustom1_GotFocus(sender As Object, e As EventArgs) Handles txtCustom1.GotFocus, txtCustom2.GotFocus, txtCustom3.GotFocus, txtCustom4.GotFocus, txtCustom5.GotFocus, txtCustom6.GotFocus, txtCustom7.GotFocus, txtCustom8.GotFocus, txtCustom9.GotFocus, txtCustom10.GotFocus, txtCustom11.GotFocus, txtCustom12.GotFocus, txtCustom13.GotFocus, txtCustom14.GotFocus, txtCustom15.GotFocus, txtRefill1.GotFocus, txtRefill2.GotFocus, txtRefill3.GotFocus, txtRefill4.GotFocus
        tmSelectAll.Enabled = False
        tmSelectAll.Tag = sender
        tmSelectAll.Enabled = True
    End Sub
    Private Sub tmSelectAll_Tick(sender As Object, e As EventArgs) Handles tmSelectAll.Tick
        tmSelectAll.Enabled = False
        Try
            tmSelectAll.Tag.SelectAll
        Catch ex As Exception
            'Pass
        End Try
    End Sub

End Class
