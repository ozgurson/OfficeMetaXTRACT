Imports System.IO
Imports System.Text.RegularExpressions
Module Module1
    Public curdir = Directory.GetCurrentDirectory
    Public ologfilename = Now.ToString("yyyyMMddHHmmss") + ".XML"
    Public logfilename As String = ""
    Function Cleanstring(dirtystring As String)
        Cleanstring = Regex.Replace(dirtystring, "[^A-Za-z0-9\-/\._]", "")
    End Function
    Function LogStart(ByVal pfilename As String, ByVal plogtype As String) As String
        Try
            Dim lf = My.Computer.FileSystem.OpenTextFileWriter(pfilename, System.IO.FileMode.CreateNew)
            Select Case plogtype
                Case = "XML"
                    lf.WriteLine("<?xml version='1.0' encoding='UTF-8'?>")
                    lf.WriteLine("<Files>")
                    lf.Close()
                Case = "CSV"
                Case = "TXT"
                Case Else
                    lf.Close()
                    Console.WriteLine("Error : " + Err.Description + vbCrLf + "LOGSTART - Unknown parameter1.")
                    GC.Collect()
                    End
            End Select
            lf.Close()
        Catch
            Console.WriteLine("Error : " + Err.Description + vbCrLf + "LOGSTART - Unknown error caught.")
            GC.Collect()
            End
        End Try
    End Function
    Function LogEnd(ByVal pfilename As String, ByVal plogtype As String) As String
        Try
            Dim lf = My.Computer.FileSystem.OpenTextFileWriter(pfilename, System.IO.FileMode.Append)
            Select Case plogtype
                Case = "XML"
                    lf.WriteLine("</Files>")
                    lf.Close()
                Case = "CSV"
                Case = "TXT"
                Case Else
                    lf.Close()
                    Return "ERROR"
            End Select
            lf.Close()
        Catch
            Return "ERROR"
        End Try
    End Function
    Function LogSectionStartEnd(ByVal pfilename As String, ByVal psectionname As String, ByVal paction As String) As String
        Try
            Dim lf = My.Computer.FileSystem.OpenTextFileWriter(pfilename, System.IO.FileMode.Append)
            'Call only log file type is XML, each file is section.
            'SectionName is filename
            Select Case paction
                Case = "START"
                    lf.WriteLine("<" + psectionname + ">")
                    lf.Close()
                Case = "END"
                    lf.WriteLine("</" + psectionname + ">")
                    lf.Close()
                Case Else
                    Return "ERROR"
            End Select
        Catch
            Return "ERROR"
        End Try
    End Function
    Function LogWrite(ByVal pfilename As String, ByVal plogtype As String, ByVal plogsource As String, ByVal ploglabel As String, ByVal plogvalue As String)
        'LOGTYPE[XML,CSV,TXT] - LOGSOURCE[WIN,XLS,DOC,PPT....] - 
        'Dim linenumber As String
        Try
            Dim lf = My.Computer.FileSystem.OpenTextFileWriter(pfilename, System.IO.FileMode.Append)
            Select Case plogtype
                Case Is = "XML"
                    lf.WriteLine("<" + ploglabel + ">" + plogvalue + "</" + ploglabel + ">")
                    lf.Close()
                Case Else
                    lf.Close()
                    Return "ERROR"
            End Select
        Catch
            Return "ERROR"
        End Try
    End Function
    Sub Main(args As String())
        If args.Length <> 2 Then
            Console.WriteLine("This program extracts filesystem and metadata information from MicroSoft Office files in the current directory")
            Console.WriteLine("Usage : <executablename> Parameter1 Parameter2")

            Console.WriteLine("Parameter1 can be any of : XML CSV TXT")
            Console.WriteLine("Parameter2 can be any of : DOC XLS PPT ALL")
            Console.WriteLine("1.[ALL extracts metadata from all three types of file (DOC,XLS,PPT)]")
            Console.WriteLine("2.[Naming for output file/files is:  XXX_YYYYMMDDHHMISS.QQQ] where XXX is one of (DOC,XLS,PPT) and YYY is ome of (XML,CSV,TXT) ")
            Console.WriteLine("  Where YYYYMMDDHHMISS stands for year month day[of month] hour minute and seconds of run time.")
            End
        End If
        args(0) = args(0).ToUpper
        args(1) = args(1).ToUpper

        If args(0) = "XML" And args(1) = "DOC" Then
            logfilename = "DOC_" + ologfilename
            Call XMLDOC()
        End If
        If args(0) = "XML" And args(1) = "XLS" Then
            logfilename = "XLS_" + ologfilename
            Call XMLXLS()
        End If
        If args(0) = "XML" And args(1) = "PPT" Then
            logfilename = "PPT_" + ologfilename
            Call XMLPPT()
        End If
        If args(0) = "XML" And args(1) = "ALL" Then
            logfilename = "DOC_" + ologfilename
            Call XMLDOC()
            logfilename = "XLS_" + ologfilename
            Call XMLXLS()
            logfilename = "PPT_" + ologfilename
            Call XMLPPT()
        End If
    End Sub
    Sub XMLDOC()
        Dim filecnt As Integer
        Dim fsinfo As FileInfo
        Dim filelist As String()
        Dim Wapp As Microsoft.Office.Interop.Word.Application
        Dim docOffice As Microsoft.Office.Interop.Word.Document
        Try
            Wapp = New Microsoft.Office.Interop.Word.Application
            Wapp.Visible = False
            filelist = System.IO.Directory.GetFiles(curdir, "*.doc?")
        Catch
            Console.WriteLine("Error : " + Err.Description + vbCrLf + "XMLDOC - Unknown error caught.[WAPP set]")
            GC.Collect()
            End
        End Try

        If filelist.Count <> 0 Then
            LogStart(logfilename, "XML")
            Console.WriteLine("")
        Else
            Console.WriteLine("Nothing to do! [ No files]")
            Wapp.Quit()
            End
        End If

        For Each file In filelist
            filecnt += 1
            Console.CursorLeft = 1
            Console.Write("Microsoft Word Documents       : Processing file " + filecnt.ToString + " of " + filelist.Count.ToString)

            fsinfo = My.Computer.FileSystem.GetFileInfo(file)
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "START")
            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "START")
            LogWrite(logfilename, "XML", "WIN", "Fullname", fsinfo.FullName)
            LogWrite(logfilename, "XML", "WIN", "Name", fsinfo.Name)
            LogWrite(logfilename, "XML", "WIN", "Length", fsinfo.Length.ToString)
            LogWrite(logfilename, "XML", "WIN", "Extension", fsinfo.Extension)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time", fsinfo.CreationTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time_UTC", fsinfo.CreationTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Directory_Name", fsinfo.DirectoryName)
            LogWrite(logfilename, "XML", "WIN", "Exists", fsinfo.Exists.ToString)
            LogWrite(logfilename, "XML", "WIN", "Is_ReadOnly", fsinfo.IsReadOnly.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time", fsinfo.LastAccessTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time_UTC", fsinfo.LastAccessTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time", fsinfo.LastWriteTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time_UTC", fsinfo.LastWriteTimeUtc.ToString)

            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "END")
            LogSectionStartEnd(logfilename, "SELF_DATA", "START")

            docOffice = Wapp.Documents.Open(file)
            Dim _BuiltInProperties As Object = docOffice.BuiltInDocumentProperties

            If _BuiltInProperties IsNot Nothing Then
                Try
                    Dim prop As Object, _label As String, _value As String
                    For Each prop In _BuiltInProperties
                        _value = "" : _label = ""
                        _label = Cleanstring(Replace(prop.Name, " ", "_"))
                        Try
                            If Not IsNothing(prop.value) And Len(prop.value) <> 0 Then
                                _value = Replace(Replace(Replace(prop.value, vbCrLf, ""), vbLf, ""), vbCr, "")
                            Else
                                _value = "EMPTY"
                            End If
                        Catch
                            _value = "NULL"
                        End Try
                        'Console.WriteLine(_label + ":" + _value)
                        LogWrite(logfilename, "XML", "DOC", _label, _value)
                    Next
                    docOffice.Close()
                Catch
                    docOffice.Close()
                End Try
            End If
            LogSectionStartEnd(logfilename, "SELF_DATA", "END")
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "END")
        Next
        LogEnd(logfilename, "XML")
        docOffice = Nothing
        Wapp.Quit()

    End Sub
    Sub XMLXLS()
        Dim filecnt As Integer
        Dim fsinfo As FileInfo
        Dim filelist As String()
        Dim Wapp As Microsoft.Office.Interop.Excel.Application
        Dim docOffice As Microsoft.Office.Interop.Excel.Workbook

        Try

            Wapp = New Microsoft.Office.Interop.Excel.Application
            Wapp.Visible = False
            filelist = System.IO.Directory.GetFiles(curdir, "*.xls?")
        Catch
            Console.WriteLine("Error : " + Err.Description + vbCrLf + "XMLXLS - Unknown error caught.[WAPP set]")
            GC.Collect()
            End
        End Try

        If filelist.Count <> 0 Then
            LogStart(logfilename, "XML")
            Console.WriteLine("")
        Else
            Console.WriteLine("Nothing to do! [ No files]")
            Wapp.Quit()
            End
        End If

        For Each file In filelist
            filecnt += 1
            Console.CursorLeft = 1
            Console.Write("Microsoft Excel Documents      : Processing file " + filecnt.ToString + " of " + filelist.Count.ToString)

            fsinfo = My.Computer.FileSystem.GetFileInfo(file)
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "START")
            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "START")
            LogWrite(logfilename, "XML", "WIN", "Fullname", fsinfo.FullName)
            LogWrite(logfilename, "XML", "WIN", "Name", fsinfo.Name)
            LogWrite(logfilename, "XML", "WIN", "Length", fsinfo.Length.ToString)
            LogWrite(logfilename, "XML", "WIN", "Extension", fsinfo.Extension)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time", fsinfo.CreationTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time_UTC", fsinfo.CreationTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Directory_Name", fsinfo.DirectoryName)
            LogWrite(logfilename, "XML", "WIN", "Exists", fsinfo.Exists.ToString)
            LogWrite(logfilename, "XML", "WIN", "Is_ReadOnly", fsinfo.IsReadOnly.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time", fsinfo.LastAccessTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time_UTC", fsinfo.LastAccessTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time", fsinfo.LastWriteTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time_UTC", fsinfo.LastWriteTimeUtc.ToString)
            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "END")
            LogSectionStartEnd(logfilename, "SELF_DATA", "START")

            docOffice = Wapp.Workbooks.Open(file)
            Dim _BuiltInProperties As Object = docOffice.BuiltinDocumentProperties

            If _BuiltInProperties IsNot Nothing Then
                Try
                    Dim prop As Object, _label As String, _value As String
                    For Each prop In _BuiltInProperties
                        _value = "" : _label = ""
                        _label = Cleanstring(Replace(prop.Name, " ", "_"))
                        Try
                            If Not IsNothing(prop.value) And Len(prop.value) <> 0 Then
                                _value = Replace(Replace(Replace(prop.value, vbCrLf, ""), vbLf, ""), vbCr, "")
                            Else
                                _value = "EMPTY"
                            End If
                        Catch
                            _value = "NULL"
                        End Try
                        'Console.WriteLine(_label + ":" + _value)
                        LogWrite(logfilename, "XML", "XLS", _label, _value)
                    Next
                    docOffice.Close()
                Catch
                    docOffice.Close()
                End Try
            End If
            LogSectionStartEnd(logfilename, "SELF_DATA", "END")
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "END")
        Next
        LogEnd(logfilename, "XML")
        docOffice = Nothing
        Wapp.Quit()

    End Sub
    Sub XMLPPT()
        Dim filecnt As Integer
        Dim fsinfo As FileInfo
        Dim filelist As String()
        Dim Wapp As Microsoft.Office.Interop.PowerPoint.Application
        Dim docOffice As Microsoft.Office.Interop.PowerPoint.Presentation
        Try
            Wapp = New Microsoft.Office.Interop.PowerPoint.Application With {.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized}
            'Wapp.Visible = False 'PPT does not support visible = false
            filelist = System.IO.Directory.GetFiles(curdir, "*.ppt?")
        Catch
            Console.WriteLine("Error : " + Err.Description + vbCrLf + "XMLPPT - Unknown error caught.[WAPP set]")
            GC.Collect()
            End
        End Try

        If filelist.Count <> 0 Then
            LogStart(logfilename, "XML")
            Console.WriteLine("")
        Else
            Console.WriteLine("Nothing to do! [ No files]")
            Wapp.Quit()
            End
        End If


        For Each file In filelist
            filecnt += 1
            Console.CursorLeft = 1
            Console.Write("Microsoft Powerpoint Documents : Processing file " + filecnt.ToString + " of " + filelist.Count.ToString)

            fsinfo = My.Computer.FileSystem.GetFileInfo(file)
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "START")
            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "START")
            LogWrite(logfilename, "XML", "WIN", "Fullname", fsinfo.FullName)
            LogWrite(logfilename, "XML", "WIN", "Name", fsinfo.Name)
            LogWrite(logfilename, "XML", "WIN", "Length", fsinfo.Length.ToString)
            LogWrite(logfilename, "XML", "WIN", "Extension", fsinfo.Extension)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time", fsinfo.CreationTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Creation_Time_UTC", fsinfo.CreationTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Directory_Name", fsinfo.DirectoryName)
            LogWrite(logfilename, "XML", "WIN", "Exists", fsinfo.Exists.ToString)
            LogWrite(logfilename, "XML", "WIN", "Is_ReadOnly", fsinfo.IsReadOnly.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time", fsinfo.LastAccessTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Access_Time_UTC", fsinfo.LastAccessTimeUtc.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time", fsinfo.LastWriteTime.ToString)
            LogWrite(logfilename, "XML", "WIN", "Last_Write_Time_UTC", fsinfo.LastWriteTimeUtc.ToString)

            LogSectionStartEnd(logfilename, "WINDOWS_DATA", "END")
            LogSectionStartEnd(logfilename, "SELF_DATA", "START")

            docOffice = Wapp.Presentations.Open(file)
            Dim _BuiltInProperties As Object = docOffice.BuiltInDocumentProperties

            If _BuiltInProperties IsNot Nothing Then
                Try
                    Dim prop As Object, _label As String, _value As String
                    For Each prop In _BuiltInProperties
                        _value = "" : _label = ""
                        _label = Cleanstring(Replace(prop.Name, " ", "_"))
                        Try
                            If Not IsNothing(prop.value) And Len(prop.value) <> 0 Then
                                _value = Replace(Replace(Replace(prop.value, vbCrLf, ""), vbLf, ""), vbCr, "")
                            Else
                                _value = "EMPTY"
                            End If
                        Catch
                            _value = "NULL"
                        End Try
                        'Console.WriteLine(_label + ":" + _value)
                        LogWrite(logfilename, "XML", "XLS", _label, _value)
                    Next
                    docOffice.Close()
                Catch
                    docOffice.Close()
                End Try
            End If
            LogSectionStartEnd(logfilename, "SELF_DATA", "END")
            LogSectionStartEnd(logfilename, "File_" + Cleanstring(fsinfo.Name), "END")
        Next
        LogEnd(logfilename, "XML")
        docOffice = Nothing
        GC.Collect()
        Wapp.Quit()

    End Sub

End Module
