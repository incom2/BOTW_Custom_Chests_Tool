Public Class clsMap

    Private Lines() As String



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Function to load the map.yml file into the string array.
    ''' </summary>
    ''' <param name="path">Path+filename</param>
    ''' <returns>TRUE if OK, FALSE if ERROR.</returns>
    ''' -----------------------------------------------------------------------
    Public Function Load(path As String) As Boolean
        Try
            Lines = IO.File.ReadAllLines(path)
            Load = True
        Catch ex As Exception
            MsgBox("Unable to load the map file:" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & ex.Message, vbExclamation, "Error")
            Load = False
        End Try
    End Function



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Function to save the string array into the map.yml file.
    ''' </summary>
    ''' <param name="path">Path+filename</param>
    ''' <returns>TRUE if OK, FALSE if ERROR.</returns>
    ''' -----------------------------------------------------------------------
    Public Function Save(path As String) As Boolean
        Try
            IO.File.WriteAllLines(path, Lines)
            Save = True
            Threading.Thread.Sleep(2000)
        Catch ex As Exception
            MsgBox("Unable to save the map file:" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & ex.Message, vbExclamation, "Error")
            Save = False
        End Try
    End Function



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Returns the editable value from a certain "DropActor" line.
    ''' Any other line can be asked but optional parameter must be FALSE.
    ''' - If line is not found, "ERROR" is returned.
    ''' - If asked a DropActor line and it's not, "ERROR" is returned.
    ''' </summary>
    ''' <param name="num">line number</param>
    ''' <param name="mustBeDropActor">Line must be a DropActor one. </param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------
    Public Function GetLine(num As Integer, Optional mustBeDropActor As Boolean = True) As String
        Dim s As String
        'First array position is 0, but first file line (in an editor) is 1.
        Try
            s = Lines(num - 1)
        Catch ex As Exception
            Return "ERROR"
        End Try
        If mustBeDropActor Then
            If Not s.StartsWith("    - '!Parameters': {DropActor: ") AndAlso Not s.EndsWith(", EnableRevival: false, IsHardModeActor: false, IsInGround: false, SharpWeaponJudgeType: 0}") Then
                Return "ERROR"
            End If
        End If
        '    - '!Parameters': {DropActor: CUSTOM_01, EnableRevival: false, IsHardModeActor: false, IsInGround: false, SharpWeaponJudgeType: 0}
        s = Replace(s, "    - '!Parameters': {DropActor: ", "")
        s = Replace(s, ", EnableRevival: false, IsHardModeActor: false, IsInGround: false, SharpWeaponJudgeType: 0}", "")
        Return s
    End Function



    ''' -----------------------------------------------------------------------
    ''' <summary>
    ''' Sets the passed value in the editable part of a certain line.
    ''' Non allowed characters, etc. will be scrapped here.
    ''' Returns TRUE if OK, FALSE if error.
    ''' </summary>
    ''' <param name="num"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------
    Public Function SetLine(num As Integer, value As String) As Boolean

        For i As Integer = 1 To Len(value)
            Dim oneChar = Strings.Mid(value, i, 1)
            Dim itsGood As Boolean = False

            If oneChar = "_" Then
                itsGood = True 'Underscore is allowed.
            Else
                Select Case Asc(oneChar)
                    Case 65 To 90, 97 To 122, 48 To 57
                        itsGood = True 'Letters (upper case, lower case) and numbers are allowed.
                End Select
            End If

            If Not itsGood Then Return False 'Bad string detected!
        Next

        'New value is OK.
        Dim newLine As String = "    - '!Parameters': {DropActor: " & value & ", EnableRevival: false, IsHardModeActor: false, IsInGround: false, SharpWeaponJudgeType: 0}"
        Lines(num - 1) = newLine
        Return True
    End Function

End Class
