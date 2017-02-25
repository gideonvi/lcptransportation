Imports System.Data
Imports System.Configuration
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Compression
Imports System.Web
Imports System.Net
Imports System.Text
Imports System.Math
Imports System.Data.SqlClient
'Imports System.Web.HttpApplication
'Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Excel = Microsoft.Office.Interop.Excel
Module ExProcess
    Dim G_LatLon(3)
    Dim G_TimeDist(3)
    Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Integer)
    Sub Main()
        '   SendTextMsg("(317)440-6212", "Please confirm your pickup for 9:15 from 123 Main Street", "Goto Testit showdowninnaptown.com/responsecenter.aspx")
        '  SendTextMsg("(970)214-0581", "Please confirm your pickup for 9:15 from 123 Main Street", "Goto Testit showdowninnaptown.com/responsecenter.aspx")
        ' SendTextMsg("(317)362-5363", "Please confirm your pickup for 9:15 from 123 Main Street", "Goto Testit showdowninnaptown.com/responsecenter.aspx")


        'SetAddress_LatLon()
        'SetDriver_LatLon()
        'SetTravelTime()
        'Threading.Thread.Sleep(10000)
        'Loop
        ' SetTripEvents()
        'SetZipDistance()
        '   End
        ' GetVichilceLocation()
        LoadCareSourceMember()
           MsgBox("test completed")
    End Sub
    Sub LoadCareSourceMember()

        Dim WrkDir

        Dim wrksql As String
        '  Try


        Console.Write("Checking for CareSoure Eligility Files" + Chr(10))
      
        WrkDir = System.Configuration.ConfigurationManager.AppSettings("EligilityFilesCareSource").ToString


        Dim picList As String() = Directory.GetFiles(WrkDir, "LCP_*.txt")
        For Each f As String In picList

            Console.Write("Processing CareSoure Eligility File:" + f + Chr(10))
            ' Create an instance of StreamReader to read from a file. 
            Dim sr As StreamReader = New StreamReader(f)
            Dim StrWer As StreamWriter
            Dim TempFile
            Dim line As String, Outrec As String
            Outrec = ""
            ' Read and display the lines from the file until the end  
            ' of the file is reached. 
            Do
                line = sr.ReadLine()
                Outrec += line

            Loop Until line Is Nothing
            sr.Close()
            TempFile = WrkDir + "caresourceraw.txt"
            StrWer = File.CreateText(TempFile)
            Outrec = Replace(Outrec, "~", vbCrLf)
            StrWer.WriteLine(Outrec)
            StrWer.Close()
            StrWer.Dispose()
            LoadCareSourceMember_upload(TempFile)
            File.GetLastWriteTime(f)
            CallProcedure("[gideon].[spMemberImport]", "LCP_ProcStrconnection", f, File.GetLastWriteTime(f), "", "")

        Next


        'Catch ex As Exception
        '    MsgBox("Load CareSource Member failed:" + ex.Message)
        'End Try

    End Sub

    Private Sub LoadCareSourceMember_upload(PassFileName)
        Dim sb As New StringBuilder
        '    Dim fd As OpenFileDialog = New OpenFileDialog()
        '  Dim fReader As StreamReader
        '  Dim FileLogWriter As StreamWriter
        ' Dim fWriter As StreamWriter
        Dim sReader As String = ""
        Dim strFileName As String
        Dim wrksql As String

        'fd.Title = "Open File Dialog"

        'fd.InitialDirectory = "T:\Reports\Care Source\MemberFiles"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        'fd.FilterIndex = 2
        'fd.RestoreDirectory = True
        Dim lineparse() As String

        Dim ISA6 As String = ""
        Dim ISA8 As String = ""
        Dim ISA9 As String = ""
        Dim ISA10 As String = ""
        Dim ISA15 As String = ""
        Dim ISA16 As String = ""

        Dim BGN2 As String = ""
        Dim BGN8 As String = ""

        Dim FNAME As String = ""
        Dim LNAME As String = ""
        Dim MNAME As String = ""
        Dim PHONE As String = ""

        Dim SPN2 As String = ""
        Dim SPN4 As String = ""

        Dim REFNO As String = ""
        Dim LOB As String = ""
        Dim MEDICAIDID As String = ""
        Dim MEDICAREID As String = ""
        Dim DOB As String = ""
        Dim Addr As String = ""
        Dim DOB2 As String = ""
        Dim CITY As String = ""
        Dim STATE As String = ""
        Dim ZIP As String = ""
        Dim GENDER As String = ""
        Dim SDATE As String = ""
        Dim EDATE As String = ""
        Dim SDATE2 As String = ""
        Dim EDATE2 As String = ""
        Dim PID As String = ""
        Dim StateID As String = "15"
        Dim MCOID As String = "22"
        Dim IsActive As String = "1"
        Dim wrkmsg



        Dim YYYY As String = ""
        Dim MM As String = ""
        Dim DD As String = ""

        wrksql = "Delete from [gideon].[TblMemberStage1]"
        TransactionIO(wrksql, "DELETE", "LCP")

        'Dim intCnt As Integer


        '  If fd.ShowDialog() = DialogResult.OK Then
        ' strFileName = fd.FileName
        strFileName = PassFileName
        ' fReader = My.Computer.FileSystem.OpenTextFileReader(strFileName)
        '  fWriter = My.Computer.FileSystem.OpenTextFileWriter(strFileName + ".out", False)
        '  FileLogWriter = My.Computer.FileSystem.OpenTextFileWriter("E:\LCP_Transportation\FileLog.txt", False)

        Dim fReader As StreamReader = New StreamReader(strFileName)
        Do While Not fReader.EndOfStream
            Try
                sReader = fReader.ReadLine()
                lineparse = sReader.Split("*")

                Select Case lineparse(0)
                    Case Is = "ISA"
                        ISA6 = Trim(lineparse(6))
                        ISA8 = Trim(lineparse(8))
                        ISA9 = Trim(lineparse(9))
                        ISA10 = Trim(lineparse(10))
                        ISA15 = Trim(lineparse(15))
                        ISA16 = Trim(lineparse(16))
                    Case Is = "BGN"
                        BGN2 = Trim(lineparse(2))
                        BGN8 = Trim(lineparse(8))
                    Case Is = "N1"
                        Select Case lineparse(1)
                            Case Is = "P5"
                                SPN2 = lineparse(2)
                            Case Is = "IN"
                                SPN4 = lineparse(4)
                        End Select
                    Case Is = "NM1"
                        Select Case lineparse(1)
                            Case Is = "IL"
                                LNAME = lineparse(3)
                                FNAME = lineparse(4)
                                MNAME = lineparse(5)
                        End Select
                    Case Is = "REF"
                        Select Case lineparse(1)
                            Case Is = "0F"
                                REFNO = lineparse(2)
                            Case Is = "1L"
                                LOB = lineparse(2)
                            Case Is = "23"
                                MEDICAIDID = lineparse(2)
                            Case Is = "3H"
                                MEDICAREID = lineparse(2)
                            Case Is = "PID"
                                PID = lineparse(2)
                        End Select
                    Case Is = "PER"
                        PHONE = lineparse(4)
                    Case Is = "DMG"
                        DOB = lineparse(2)
                        YYYY = Mid(DOB, 1, 4)
                        MM = Mid(DOB, 5, 2)
                        DD = Mid(DOB, 7, 2)
                        DOB2 = DateSerial(YYYY, MM, DD)
                        GENDER = lineparse(3)
                    Case Is = "N3"
                        Addr = lineparse(1)
                    Case Is = "N4"
                        CITY = lineparse(1)
                        STATE = lineparse(2)
                        ZIP = lineparse(3)
                    Case Is = "DTP"
                        Select Case lineparse(1)
                            Case Is = "348"
                                SDATE = lineparse(3)
                                YYYY = Mid(SDATE, 1, 4)
                                MM = Mid(SDATE, 5, 2)
                                DD = Mid(SDATE, 7, 2)
                                SDATE2 = DateSerial(YYYY, MM, DD)
                            Case Is = "349"
                                EDATE = lineparse(3)
                                YYYY = Mid(EDATE, 1, 4)
                                MM = Mid(EDATE, 5, 2)
                                DD = Mid(EDATE, 7, 2)
                                EDATE2 = DateSerial(YYYY, MM, DD)
                        End Select
                    Case Is = "LX"
                        ' fWriter.Write(MEDICAIDID & "," & MEDICAREID & "," & FNAME & "," & MNAME & "," & LNAME & "," & PHONE & "," & Addr & "," & CITY & "," & STATE & "," & ZIP & "," & GENDER & "," & DOB2 & "," & SDATE2 & "," & EDATE2 & "," & "15" & "," & "22" & "," & "1" & "," & PID & "," & LOB & vbCrLf)
                        wrksql = " INSERT INTO [gideon].[TblMemberStage1]"
                        wrksql += "            ([MedicaidID]"
                        wrksql += "            ,[MedicareID]"
                        wrksql += "            ,[FirstName]"
                        wrksql += "            ,[MiddleName]"
                        wrksql += "            ,[LastName]"
                        wrksql += "            ,[Phone]"
                        wrksql += "            ,[Addr1]"
                        wrksql += "            ,[City]"
                        wrksql += "            ,[State]"
                        wrksql += "            ,[ZipCode]"
                        wrksql += "            ,[Sex]"
                        wrksql += "            ,[DateOfBirth]"
                        wrksql += "            ,[EligibilityStartDate]"
                        wrksql += "            ,[EligibilityEndDate]"
                        wrksql += "            ,[StateID]"
                        wrksql += "            ,[MCOID]"
                        wrksql += "            ,[IsActive]"
                        wrksql += "            ,[PID]"
                        wrksql += "           )"
                        wrksql += "      VALUES"
                        wrksql += "            ('" + MEDICAIDID + "'"
                        wrksql += "            ,'" + MEDICAREID + " '"
                        wrksql += "            ,'" + FNAME + "'"
                        wrksql += "            ,'" + MNAME + "'"
                        wrksql += "            ,'" + LNAME + "'"
                        wrksql += "            ,'" + PHONE + "'"
                        wrksql += "            ,'" + Addr + "'"
                        wrksql += "            ,'" + CITY + "'"
                        wrksql += "            ,'" + STATE + "'"
                        wrksql += "            ,'" + ZIP + "'"
                        wrksql += "            ,'" + GENDER + "'"
                        wrksql += "            ,'" + DOB2 + "'"
                        wrksql += "            ,'" + SDATE2 + "'"
                        wrksql += "            ,'" + EDATE2 + "'"
                        wrksql += "            ,'" + StateID + "'"
                        wrksql += "            ,'" + MCOID + "'"
                        wrksql += "            ,'" + IsActive + "'"
                        wrksql += "            ,'" + PID + " '"
                        wrksql += "            "
                        wrksql += ")"
                        wrkmsg = TransactionIO(wrksql, "INSERT", "LCP")
                        If Len(wrkmsg) > 3 Then
                            wrkmsg = "Insert Error:" + wrkmsg
                            ' SendProcessError("Load CareSource Member", wrkmsg)
                            Exit Sub
                        End If
                End Select

            Catch ex As Exception
                ' SendProcessError("Load CareSource Member", ex.ToString)
            End Try
        Loop
        '
        fReader.Close()
        fReader.Dispose()
    End Sub

    Sub SetTripEvents()
        Dim lConnection As New OleDbConnection("")
        Dim lCommand As OleDbCommand
        Dim lDataReader As OleDbDataReader
        '
        Dim WrkCount1, WrkCount2
        Dim wrkEvents(5, 3)

        wrkEvents(1, 1) = "PPC"
        wrkEvents(1, 2) = "Event PPC (Passenger Pickup Confirmation)"
        wrkEvents(2, 1) = "PPL"
        wrkEvents(2, 2) = "Event PPL (Event Passenger Pickup Request Link)"
        Dim lclConnect As New SqlConnection, lclCommand As New SqlCommand

        Dim wrkConnection = GetAppSetting("LCP_ProcStrconnection")
        lclConnect = New SqlConnection(wrkConnection)
        lclConnect.Open()
        lclCommand.Connection = lclConnect
        'Event Passenger Pickup Request 
        For E1 = 1 To 2
            lclCommand.Parameters.Clear()
            Console.Write(wrkEvents(E1, 2) + Chr(10))
            WrkCount1 = Dcount("Event_ID", "Gideon.TblEvent_Control_Detail", "", "LCP")
            lclCommand.CommandText = "[gideon].[spEventCreation]"
            lclCommand.CommandType = CommandType.StoredProcedure
            lclCommand.Parameters.AddWithValue("@PassAction", wrkEvents(E1, 1))
            lclCommand.Parameters.AddWithValue("@PassTripID", 0)
            lclCommand.Parameters.AddWithValue("@PassTripLeg", 0)
            lclCommand.Parameters.AddWithValue("@PassFromStatus", 0)
            lclCommand.Parameters.AddWithValue("@PassToStatus", 0)
            lclCommand.ExecuteNonQuery()
            WrkCount2 = Dcount("Event_ID", "Gideon.TblEvent_Control_Detail", "", "LCP")
            Console.Write("Event PPC Added " + Str(WrkCount2 - WrkCount1) + Chr(10))
        Next
       
        lclConnect.Close()

    End Sub
    Function GetVichilceLocation()
        Dim url As String = " https://auth.networkfleet.com/token"
        '`Headers-->
        '`Authorization: Basic czZCaGRSa3F0MzpnWDFmQmF0M2JW
        '`Content-Type: application/x-www-form-urlencoded

        '`Request body-->
        '`grant_type=password&username=yourClientId&password=yourClientSecret
        url += "?grant_type=password&username=LCPMileEX&password=LCPMikeEX"

        Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
        Dim receiveStream As Stream = myHttpWebResponse.GetResponseStream()
        Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim readStream As New StreamReader(receiveStream, encode)
        '   Console.WriteLine(ControlChars.Lf + ControlChars.Cr + "Response stream received")
        Dim read(10200) As [Char]
        Dim count As Integer = readStream.Read(read, 0, 10200)
        Dim str As New [String](read, 0, count)
       
        ' Releases the resources of the Stream.

        Dim wrkToken, x, y, r, wrkRefreshToken

        '"refresh_token":"
         
        x = InStr(str, ":")
        y = InStr(str, ",")
        If x > 0 And y > x Then
            wrkToken = Mid(str, x + 2, y - x - 3)
        Else
            Exit Function
        End If

        r = InStr(str, "refresh_token")
        x = InStr(r, str, ":")
        y = InStr(r, str, ",")
        If x > 0 And y > x Then
           wrkRefreshToken= Mid(str, x + 2, y - x - 3)
        Else
            Exit Function
        End If
        readStream.Close()
        readStream.Dispose()
        '  MsgBox(wrkToken + Chr(10) + str)

        'url = " GET /locations/vehicle/{vehicleId}/track?{queryParam1}&{queryParam2}... HTTP/1.1"
        'url += " Host:           https://api.networkfleet.com"
        'url += " Content-Type:  [application/vnd.networkfleet.api-v1+json] [application/vnd.networkfleet.api-v1+xml] "
        'url += " Accept:        [application/vnd.networkfleet.api-v1+json] [application/vnd.networkfleet.api-v1+xml] "
        'url += " Authorization: Bearer {access-token}"
        'url += " Cache-Control: no-cache"
        'url = Replace(url, "access-token", wrkToken)
        '
        url = "https://api.networkfleet.com/"
        url = "https://api.networkfleet.com/locations/vehicle/"
        Dim myHttpWebRequest2 As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)

        '  myHttpWebRequest2.Method = "GET"
        '    myHttpWebRequest2.Host = "https://api.networkfleet.com/locations/vehicle/"

        myHttpWebRequest2.Headers("Authorization") = " " & wrkRefreshToken & " "
        myHttpWebRequest2.ContentType = "[application/vnd.networkfleet.api-v1+json] [application/vnd.networkfleet.api-v1+xml]"
        myHttpWebRequest2.Accept = " [application/vnd.networkfleet.api-v1+json] [application/vnd.networkfleet.api-v1+xml]"
        myHttpWebRequest2.CachePolicy = Nothing


        Dim myHttpWebResponse2 As HttpWebResponse = CType(myHttpWebRequest2.GetResponse(), HttpWebResponse)
        Dim receiveStream2 As Stream = myHttpWebResponse.GetResponseStream()
        Dim encode2 As Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim readStream2 As New StreamReader(receiveStream2, encode)
        '   Console.WriteLine(ControlChars.Lf + ControlChars.Cr + "Response stream received")
        Dim read2(10200) As [Char]
        Dim count2 As Integer = readStream.Read(read2, 0, 10200)
        Dim str2 As New [String](read, 0, count)

    End Function
    Sub SetZipDistance()


        Dim wrksql, wrkmsg
        Dim wrkaddr(10, 2)
        Dim wrklat(2)
        Dim wrklon(2)

        wrksql = "select  *  from [gideon].[tblZipCodeLookup] where DistanceEstimate is null or TimeEstimate is null "


        Dim lConnection2 As New OleDbConnection("")
        Dim lCommand2 As OleDbCommand
        Dim lDataReader2 As OleDbDataReader

        Dim lConnection As New OleDbConnection("")
        Dim lCommand As OleDbCommand
        Dim lDataReader As OleDbDataReader
        '
        Dim WrkStartCity, WrkEndCity
        Dim wrktime
        Console.Write("Starting Zip Code" + Chr(10))
        lConnection.ConnectionString = GetAppSetting("LCPstrconnection")
        lConnection.Open()
        lCommand = New OleDbCommand(wrksql, lConnection)
        lDataReader = lCommand.ExecuteReader
        If lDataReader.HasRows Then
            While lDataReader.Read()
                 
                WrkStartCity = lDataReader("StartZip") + "," + DLookup("StartCity", "[Mite].[TripLeg]", "StartZip = '" + lDataReader("StartZip") + "' and not StartCity is null", "LCP")
                WrkEndCity = lDataReader("EndZip") + "," + DLookup("EndCity", "[Mite].[TripLeg]", "EndZip = '" + lDataReader("EndZip") + "' and not EndCity is null", "LCP")

                Console.Write("Zip Code: " + WrkStartCity + " to " + WrkEndCity + Chr(10))

                GetTravalTimeDist(WrkStartCity, WrkEndCity)

                If Len(G_TimeDist(3)) < 3 Then
                    wrktime = Str(Val(ConvertToMins(G_TimeDist(1))))
                    wrksql = " update [gideon].[tblZipCodeLookup] set "
                    wrksql += "   DistanceEstimate= " + Str(Val(G_TimeDist(2))) + ","
                    wrksql += "   TimeEstimate= " + wrktime
                    wrksql += " where ZipLookupID =  " + Str(lDataReader("ZipLookupID"))
                    wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")
                    If Len(wrkmsg) > 0 Then
                        MsgBox(wrkmsg + Chr(10) + wrksql)
                        Exit Sub
                    End If
                Else
                    Console.Write("Error:" + G_LatLon(3))
                End If
 
            End While
        End If
        lDataReader.Close()
        lConnection.Close()

    End Sub




    Sub SetDriver_LatLon()
        Dim wrksql, wrkmsg

        Console.Write("Adding Driver Lat/Lon records" + Chr(10))
        wrksql = " insert into [gideon].[tblDriver_Lat_Lon]  "
        wrksql += " select driver_id,null,null  "
        wrksql += " from  [gideon].[vwDriver] as d1"
        wrksql += " where  not exists (select 1 from [gideon].[tblDriver_Lat_Lon] as d2"
        wrksql += "                      where d2.driver_id = d2.driver_id"
        wrksql += "                     )"
        wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")
        If Len(wrkmsg) > 0 Then
            MsgBox(wrkmsg + Chr(10) + wrksql)
            Exit Sub
        End If

        wrksql = " select  d1.driver_id,caddr1,ccity,czip"
        wrksql += "  from  [gideon].[vwDriver] as d1,"
        wrksql += "        [gideon].[tblDriver_Lat_Lon] as d2"
        wrksql += "  where d1.driver_id =  d2.Driver_ID "
        wrksql += "    and isnull(d2.Latitude,0) = 0"

        Dim lConnection2 As New OleDbConnection("")
        Dim lCommand2 As OleDbCommand
        Dim lDataReader2 As OleDbDataReader

        '
        Dim wrkaddr
        Console.Write("Starting Driver Geo Code" + Chr(10))
        lConnection2.ConnectionString = GetAppSetting("LCPstrconnection")
        lConnection2.Open()
        lCommand2 = New OleDbCommand(wrksql, lConnection2)
        lDataReader2 = lCommand2.ExecuteReader
        If lDataReader2.HasRows Then
            While lDataReader2.Read()
                ' 
                G_LatLon(1) = 0
                G_LatLon(2) = 0
                wrkaddr = lDataReader2("caddr1") + "," + lDataReader2("ccity") + "," + lDataReader2("czip")
                GetLatLon(wrkaddr)
                wrkmsg = wrkaddr
                wrkmsg += Chr(10)
                Console.Write(wrkmsg)
                If Val(G_LatLon(1)) <> 0 Then
                    wrksql = "update [gideon].[tblDriver_Lat_Lon] set [Latitude]=#4 ,[Longitude]=#5"
                    wrksql += " where driver_id = " + Str(lDataReader2("Driver_id"))
                    wrksql = Replace(wrksql, "#4", Str(Val(G_LatLon(1))))
                    wrksql = Replace(wrksql, "#5", Str(Val(G_LatLon(2))))
                    wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")
                    If Len(wrkmsg) > 0 Then
                        MsgBox(wrkmsg + Chr(10) + wrksql)
                        Exit Sub
                    End If
                End If
            End While
        End If
        lDataReader2.Close()
        lConnection2.Close()
    End Sub
    Sub SetAddress_LatLon()
        Dim wrksql, wrkmsg

        wrksql = " select [gideon].[fncExtractAddress](addr1,'F') as addrx,city,max(substring(zip,1,5)) as Zipcode,min(addr1),max(addr1),count(*)"
        wrksql += "   from ("
        wrksql += " 		select ltrim([StartAddress1]) as addr1,[StartCity] as city,[StartZip] as zip from [gideon].[tblTripLegs]"
        wrksql += " 		union all select ltrim([EndAddress1]),[EndCity],[EndZip] from [gideon].[tblTripLegs]"
        wrksql += " 	 ) as x"
        wrksql += " where not exists (select 1 from [gideon].[tblAddress_Lat_Lon] where [Address] =  gideon.[fncExtractAddress](addr1,'F'))"
        wrksql += " group by  gideon.[fncExtractAddress](addr1,'F'),city "

        Dim lConnection2 As New OleDbConnection("")
        Dim lCommand2 As OleDbCommand
        Dim lDataReader2 As OleDbDataReader

        '
        Dim wrkaddr
        Console.Write("Starting Address Geo Code" + Chr(10))
        lConnection2.ConnectionString = GetAppSetting("LCPstrconnection")
        lConnection2.Open()
        lCommand2 = New OleDbCommand(wrksql, lConnection2)
        lDataReader2 = lCommand2.ExecuteReader
        If lDataReader2.HasRows Then
            While lDataReader2.Read()
                ' 
                wrkaddr = lDataReader2("addrx") + "," + lDataReader2("city") + "," + lDataReader2("zipcode")
                GetLatLon(wrkaddr)
                wrkmsg = lDataReader2("addrx")
                wrkmsg += Chr(10)
                Console.Write(wrkmsg)
                wrksql = "INSERT INTO [gideon].[tblAddress_Lat_Lon]   ([Address] ,[City],[Zip] ,[Latitude] ,[Longitude],[ErrorMsg])"
                wrksql += "values('#1','#2','#3',#4,#5,'#6')"
                wrksql = Replace(wrksql, "#1", lDataReader2("addrx"))
                wrksql = Replace(wrksql, "#2", lDataReader2("city"))
                wrksql = Replace(wrksql, "#3", lDataReader2("zipcode"))
                wrksql = Replace(wrksql, "#4", Str(Val(G_LatLon(1))))
                wrksql = Replace(wrksql, "#5", Str(Val(G_LatLon(2))))
                wrksql = Replace(wrksql, "#6", G_LatLon(3))
                wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")
                If Len(wrkmsg) > 0 Then
                    MsgBox(wrkmsg + Chr(10) + wrksql)
                    Exit Sub
                End If
            End While
        End If
        lDataReader2.Close()
        lConnection2.Close()

        wrksql = "  update [gideon].[tblTripLegs] "
        wrksql += "    set StartLatitude =Latitude,"
        wrksql += "        startLongitude = [Longitude]"
        wrksql += " from gideon.[tblAddress_Lat_Lon]"
        wrksql += " where [Address] = gideon.[fncExtractAddress]( [StartAddress1],'F')"
        wrksql += "   and  city = [StartCity]"
        wrksql += "   and  zip = substring([Startzip],1,5)"
        wrksql += "   and  StartLatitude is null"
        wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")

        wrksql = " update [gideon].[tblTripLegs] "
        wrksql += "    set endLatitude =Latitude,"
        wrksql += "        endLongitude = [Longitude]"
        wrksql += " from gideon.[tblAddress_Lat_Lon]"
        wrksql += " where [Address] = gideon.[fncExtractAddress]( [EndAddress1],'F')"
        wrksql += "   and  city = [endCity]"
        wrksql += "    and  zip = substring([endzip],1,5)"
        wrksql += "    and endLatitude is null"
        wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")

    End Sub
    Sub SetTravelTime()
        Dim wrksql, wrkmsg, wrkcount



        wrksql = " select  top 4 * from [gideon].[tblTripLegs]"
        wrksql += " where  isnull(EstTravelTime,0) in( 0 ,-1) "
        wrksql += "   and  not endLatitude  is null  "
        wrksql += "   and  not endLongitude  is null  "
        wrksql += "   and  not startLatitude  is null  "
        wrksql += "   and  not startLongitude is null  "
    
        Dim lConnection2 As New OleDbConnection("")
        Dim lCommand2 As OleDbCommand
        Dim lDataReader2 As OleDbDataReader

        '
        Dim wrktime
        Console.Write("Calculating Travel times" + Chr(10))
        lConnection2.ConnectionString = GetAppSetting("LCPstrconnection")
        lConnection2.Open()
        lCommand2 = New OleDbCommand(wrksql, lConnection2)
        lDataReader2 = lCommand2.ExecuteReader
        If lDataReader2.HasRows Then
            While lDataReader2.Read()
                ' 
                GetTravalTimeDist(Str(lDataReader2("StartLatitude")) + "," + Str(lDataReader2("StartLongitude")), Str(lDataReader2("EndLatitude")) + "," + Str(lDataReader2("Endlongitude")))

                wrktime = G_TimeDist(1)
                Console.Write(Str(lDataReader2("TripID")) + " time :" + wrktime + Chr(10))
                wrksql = "update [gideon].[tblTripLegs] set  "
                wrksql += "  EstTravelTimeText = '" + wrktime + "',"
                wrksql += "  EstTravelTime  =  " + Str(Val(ConvertToMins(wrktime)))
                wrksql += " where [TripID] = " + Str(lDataReader2("TripID"))
                wrkmsg = TransactionIO(wrksql, "UPDATE", "LCP")
                If Len(wrkmsg) > 0 Then
                    MsgBox(wrkmsg + Chr(10) + wrksql)
                    Exit Sub
                End If
            End While
        End If
        lDataReader2.Close()
        lConnection2.Close()
    End Sub
 


    Sub GetTravalTimeDist(PassOrgin, PassDestinations)

        ' PassOrgin/PassDestinations
        '1) Lat,Lon to Lat,Lon
        '2) Full Address to Full Address
        '3) Zip to Zip
        Dim url As String = " https://maps.googleapis.com/maps/api/distancematrix/json?units=imperial&origins=#1&destinations=#2&key="
        Dim s, e, SaveStr
        Dim wrkMins, wrktime
        Dim wrkMile, wrkDist
        url = Replace(url, "#1", PassOrgin)
        url = Replace(url, "#2", PassDestinations)
        G_TimeDist(3) = ""
         Try

            Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
            Dim receiveStream As Stream = myHttpWebResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
            Dim readStream As New StreamReader(receiveStream, encode)
            '   Console.WriteLine(ControlChars.Lf + ControlChars.Cr + "Response stream received")
            Dim read(10200) As [Char]
            ' Reads 256 characters at a time.     
            Dim count As Integer = readStream.Read(read, 0, 10200)
            While count > 0
                ' Dumps the 256 characters to a string and displays the string to the console. 
                Dim str As New [String](read, 0, count)
                If InStr(str, "OVER_QUERY_LIMIT") > 0 Then
                    G_TimeDist(3) = "OVER_QUERY_LIMIT"
                    Exit While
                End If

                SaveStr = str
                s = InStr(str, "duration")
                If s > 0 Then
                    str = Mid(str, s, 200)
                    s = InStr(str, "text")
                    str = Mid(str, s, 200)
                    s = InStr(str, ":")
                    wrktime = Mid(str, s + 1, 20)
                    wrktime = Trim(Replace(wrktime, """", ""))
                    G_TimeDist(1) = wrktime
                End If

                str = SaveStr
                s = InStr(str, "distance")
                If s > 0 Then
                    str = Mid(str, s, 200)
                    s = InStr(str, "text")
                    str = Mid(str, s, 200)
                    s = InStr(str, ":")
                    wrkDist = Mid(str, s + 1, 20)
                    wrkDist = Trim(Replace(wrkDist, """", ""))
                    G_TimeDist(2) = Val(wrkDist)
                    Exit While
                End If
                '      HttpContext.Current.Response.Write(str)
                count = readStream.Read(read, 0, 256)
            End While
            ' Releases the resources of the Stream.
            readStream.Close()
            myHttpWebResponse.Close()

        Catch ex As Exception
            '   MsgBox(ex.Message)
            G_TimeDist(3) = ex.Message
        End Try

    End Sub
    Sub testit()
        MsgBox(ConvertToMins("1 hour 23 mins"))
    End Sub
    Function ConvertToMins(PassTime)
        ' "1 hour 23 mins",  
        Dim x, wrkmins
        Dim wrkvalue

        x = InStr(PassTime, "hour")
        If x > 0 Then
            wrkmins = Val(PassTime) * 60
            x += 5
            wrkmins += Val(Mid(PassTime, x, 3))
        Else
            wrkmins = Val(PassTime)
        End If
        ConvertToMins = wrkmins
    End Function
    Sub GetLatLon(ByVal PassAddr) ' As Array

        Dim url As String = "http://maps.googleapis.com/maps/api/geocode/xml?address=#1&sensor=false&key=AIzaSyAtnqEFOCn_6qUhuwh-rWnhD0euBjyG-4I"
        Dim s, e

        Try


            url = Replace(url, "#1", PassAddr)

            G_LatLon(1) = 0
            Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
            Dim receiveStream As Stream = myHttpWebResponse.GetResponseStream()
            Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
            Dim readStream As New StreamReader(receiveStream, encode)
            '  Console.WriteLine(ControlChars.Lf + ControlChars.Cr + "Response stream received")
            Dim read(10200) As [Char]
            ' Reads 256 characters at a time.     
            Dim count As Integer = readStream.Read(read, 0, 10200)
            While count > 0
                ' Dumps the 256 characters to a string and displays the string to the console. 
                Dim str As New [String](read, 0, count)
                If InStr(str, "OVER_QUERY_LIMIT") > 0 Then
                    G_LatLon(3) = "Over API limit"
                    Exit While
                End If

                s = InStr(str, "<lat>")
                e = InStr(str, "</lat>")
                If s > 0 And e > 0 Then
                    s += 5
                    G_LatLon(1) = Convert.ToDouble(Mid(str, s, e - s))
                    s = InStr(str, "<lng>")
                    e = InStr(str, "</lng>")
                    s += 5
                    G_LatLon(2) = Convert.ToDouble(Mid(str, s, e - s))
                    G_LatLon(3) = ""
                    Exit While
                End If

                '      HttpContext.Current.Response.Write(str)
                count = readStream.Read(read, 0, 256)
            End While
            ' Releases the resources of the Stream.
            readStream.Close()
            myHttpWebResponse.Close()
        Catch ex As Exception
            If InStr(ex.Message, "Length") = 0 Then
                G_LatLon(3) = "Get Lan/Lon Error: " + ex.Message
            End If

        End Try

    End Sub

    Function AddMins(ByVal PassMins, ByVal PassAddON)
        Dim wrkmins
        Dim wrkminX As String
        Dim WrkAddOn

        WrkAddOn = PassAddON
        If PassAddON > 59 Then
            WrkAddOn = Round(PassAddON / 60, 0) * 100
            WrkAddOn += PassAddON - Round(PassAddON / 60, 0)
        End If
        wrkmins = PassMins + WrkAddOn
        wrkminX = Right(CStr(wrkmins), 2)
        If Val(wrkminX) > 60 Then

        End If
        AddMins = wrkmins

    End Function
    Function SendTextMsg(PassPhoneNumber, PassSubj, PassMsg)
        Dim lConnection2 As New OleDbConnection("")
        Dim lCommand2 As OleDbCommand
        Dim lDataReader2 As OleDbDataReader
        Dim wrksql, wrkcell


        wrksql = "select * from XXCellProvider  where Active = 'Y' "

        lConnection2.ConnectionString = GetAppSetting("LCPstrconnection")
        lConnection2.Open()
        lCommand2 = New OleDbCommand(wrksql, lConnection2)
        lDataReader2 = lCommand2.ExecuteReader
        If lDataReader2.HasRows Then
            While lDataReader2.Read()
                wrkcell = FormatCellNumber(PassPhoneNumber)
                If wrkcell = "X" Then
                    Return ("Invalid Cell number")
                End If
                wrkcell += lDataReader2("PostAddress")
                SendEmail(PassMsg, wrkcell, PassSubj, "")
            End While
        End If
        lDataReader2.Close()
        lConnection2.Close()

    End Function
    Function FormatCellNumber(PassPhoneNumber)
        Dim wrkChr As String
        Dim wrkphone As String = ""

        For x = 1 To Len(PassPhoneNumber)
            If IsNumeric(Mid(PassPhoneNumber, x, 1)) Then
                wrkphone += Mid(PassPhoneNumber, x, 1)
            End If
        Next

        If Len(wrkphone) <> 10 Then
            Return "X"
        End If
        FormatCellNumber = wrkphone
    End Function
    Function GetAppSetting(ByVal PassParm)
        Try
            GetAppSetting = ConfigurationManager.AppSettings.Item(PassParm).ToString()
        Catch e As Exception
            '   WriteLog("GetAppSetting: " + e.Message, "Y")

        End Try
    End Function
    Function SendEmail(PassMsg, PassSendTo, PassSubj, PassAttachMent)

        '  Dim OutlookMessage As Outlook.MailItem
        ' Dim AppOutlook As New Outlook.Application

        'Dim objNS As Outlook._NameSpace = AppOutlook.Session
        'Dim objFolder As Outlook.MAPIFolder
        'objFolder = objNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)

        '   Try
        'OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
        'Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
        'Recipents.Add(PassSendTo)

        '        OutlookMessage.Subject = PassSubj
        '       OutlookMessage.Body = PassMsg

        'If File.Exists(PassAttachMent) Then
        'OutlookMessage.Attachments.Add(PassAttachMent)
        'End If
        'OutlookMessage.Send()

        'Catch ex As Exception
        '    '  WriteLog("DMAX: " + ex.Message + "SQL:" + wrksql, "Y")
        'Finally
        '    OutlookMessage = Nothing
        '    AppOutlook = Nothing
        'End Try

        '  Exit Function

    End Function
  




    Function TransactionIO(ByVal PassSql, ByVal PassAction, ByVal PassDb)
        Dim strconnection As String, wrkrtn
        Try

            strconnection = GetAppSetting(PassDb + "strconnection")

            Dim WrkCommand As New OleDbCommand()
            Dim WrkConnect As New OleDbConnection(strconnection)
            WrkConnect.Open()
            WrkCommand.Connection = WrkConnect
            WrkCommand.CommandText = PassSql
            wrkrtn = WrkCommand.ExecuteScalar
            WrkConnect.Close()
            wrkrtn = ""
        Catch ex As Exception
            wrkrtn = "TransactionIO:" + ex.Message
            MsgBox(wrkrtn)
        End Try
        TransactionIO = wrkrtn

    End Function
    Public Function DLookup(ByVal PassField, ByVal PassTable, ByVal PassWhere, ByVal PassDb) As String
        Dim wrksql As String
        Dim Wrkreturn


        Dim LCom_Connection As New OleDbConnection
        Dim LCom_Command As OleDbCommand
        Dim LCom_DataReader As OleDbDataReader
        Dim Connectstring As String
        Dim wrkconnection As String
        Try


            wrksql = "select " + PassField + " as ret"
            wrksql += " from " + PassTable
            If Len(PassWhere) > 0 Then
                wrksql += " where " + PassWhere
            End If


            wrkconnection = PassDb + "strconnection"


            Connectstring = GetAppSetting(wrkconnection)
            If Mid(PassDb, 1, 3) = "SQL" Then
                Connectstring = GetAppSetting(PassDb)
                LCom_Connection = New OleDbConnection(Connectstring)
                LCom_Connection.Open()
                LCom_Command = LCom_Connection.CreateCommand()
                ' lCommand = New OleDbCommand(wrksql, lConnection)
                LCom_Command.CommandText = wrksql
            Else
                LCom_Connection.ConnectionString = Connectstring
                LCom_Connection.Open()
                LCom_Command = New OleDbCommand(wrksql, LCom_Connection)
            End If


            LCom_DataReader = LCom_Command.ExecuteReader
            If LCom_DataReader.Read Then
                Wrkreturn = LCom_DataReader("ret")
            End If
            LCom_DataReader.Close()
            LCom_Connection.Close()

            If Not IsDBNull(Wrkreturn) Then
                DLookup = Wrkreturn
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function Dcount(ByVal PassField, ByVal PassTable, ByVal PassWhere, ByVal PassDb) As Integer

        Dim wrksql As String
        Dim Wrkreturn As Integer = 0
        Dim LCom_Connection As New OleDbConnection
        Dim LCom_Command As OleDbCommand
        Dim LCom_DataReader As OleDbDataReader
        Dim Connectstring As String
        Dim wrkconnection As String
        'Try
        wrksql = "select count(" + PassField + ") as ret"
        wrksql += " from " + PassTable
        If Len(PassWhere) > 0 Then
            wrksql += " where " + PassWhere
        End If


        wrkconnection = PassDb + "strconnection"


        Connectstring = GetAppSetting(wrkconnection)


        LCom_Connection.ConnectionString = Connectstring
        LCom_Connection.Open()

        LCom_Command = New OleDbCommand(wrksql, LCom_Connection)
        LCom_DataReader = LCom_Command.ExecuteReader
        If LCom_DataReader.Read Then
            If Not IsDBNull(LCom_DataReader("ret")) Then Wrkreturn = LCom_DataReader("ret")
        Else

        End If
        LCom_DataReader.Close()
        LCom_Connection.Close()
        'Catch ex As Exception
        '    MsgBox("Error DCount " + ex.Message)
        '    Try
        '        If LCom_Connection.State = ConnectionState.Open Then
        '            LCom_Connection.Close()
        '        End If
        '    Catch
        '    End Try
        'End Try
        Dcount = Wrkreturn

    End Function
    Public Function DMax(ByVal PassField, ByVal PassTable, ByVal PassWhere, ByVal PassDb) As Integer

        Dim wrksql As String
        Dim Wrkreturn As Integer = 0
        Dim LCom_Connection As New OleDbConnection
        Dim LCom_Command As OleDbCommand
        Dim LCom_DataReader As OleDbDataReader

        Try
            wrksql = "select max(" + PassField + ") as ret"
            wrksql += " from " + PassTable
            If Len(PassWhere) > 0 Then
                wrksql += " where " + PassWhere
            End If

            LCom_Connection.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(PassDb).ToString
            LCom_Connection.Open()
            LCom_Command = New OleDbCommand(wrksql, LCom_Connection)
            LCom_DataReader = LCom_Command.ExecuteReader
            If LCom_DataReader.Read Then
                If Not IsDBNull(LCom_DataReader("ret")) Then Wrkreturn = LCom_DataReader("ret")
            Else

            End If
            LCom_DataReader.Close()
            LCom_Connection.Close()
        Catch ex As Exception
            Try
                If LCom_Connection.State = ConnectionState.Open Then
                    LCom_Connection.Close()
                End If
            Catch
            End Try
        End Try
        DMax = Wrkreturn

    End Function

    Sub CallProcedure(PassProcedure, PassDb, PassParm1, PassParm2, PassParm3, PassParm4)
        Dim lConnection As New OleDbConnection("")
        '
        '   Try


        Dim lclConnect As New SqlConnection, lclCommand As New SqlCommand
        lclConnect.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(PassDb).ToString
        '   Dim wrkConnection = GetConfigSetting("LCP_ProcStrconnection")
        '   lclConnect = New SqlConnection(wrkConnection)
        lclConnect.Open()
        lclCommand.Connection = lclConnect
        lclCommand.CommandTimeout = 6000000
        lclCommand.Parameters.Clear()

        lclCommand.CommandText = PassProcedure
        lclCommand.CommandType = CommandType.StoredProcedure

        lclCommand.Parameters.Add("InFileName", SqlDbType.VarChar).Value = PassParm1

        lclCommand.Parameters.Add("INFileDateTime", SqlDbType.DateTime).Value = PassParm2

        lclCommand.ExecuteNonQuery()

        lclConnect.Close()

        'Catch ex As Exception
        '    HttpContext.Current.Response.Write("The save failed:" + ex.Message)
        'End Try





    End Sub
End Module
