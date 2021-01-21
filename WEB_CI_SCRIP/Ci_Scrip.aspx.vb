Imports System.Data.OracleClient
Imports System.Data
Imports OfficeOpenXml
Imports System.IO
Imports System.Text
Imports OfficeOpenXml.Style
Imports Ionic.Zip
Imports System.IO.Compression
Imports System.Net.Mail

Public Class Ci_Scrip
    Inherits System.Web.UI.Page

    <Obsolete>
    Public Shared Function getsqlquery(ByVal strsql As String, ByVal myconnstring As String) As DataTable
        Dim result As New DataTable
        Using connection As New OracleConnection
            connection.ConnectionString = myconnstring
            connection.Open()
            Using command As New OracleCommand(strsql, connection)
                Using reader As OracleDataReader = command.ExecuteReader()
                    result.Load(reader)
                End Using
            End Using
        End Using
        Return result
    End Function

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Label_Catch.Visible = False
        Label_ValNull.Visible = True
        Btn_Donwload.Visible = False
        If Not Me.IsPostBack Then
            Snap_Dat.Text = DateTime.Now().ToString("yyyy-MM-dd")
        End If
    End Sub
    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        Query_Scrip()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Query_Scrip()
        Label_Catch.Visible = False
    End Sub

    <Obsolete>
    Sub Query_Scrip()
        Dim Kode_Saham_Src
        Dim Kode_Emiten_Src
        Dim date_src As String
        Dim conv_date
        Dim arr_bulan() As String = {"", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "OCT", "November", "Desember"}
        date_src = ""
        If Snap_Dat.Text <> "" Then
            conv_date = Snap_Dat.Text.Split("-")
            date_src = arr_bulan(conv_date(1)) & "-" & conv_date(0).Substring(2, 2)
        End If
        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"
        Dim checker As Regex = New Regex("^[a-zA-Z0-9]*$", RegexOptions.IgnoreCase)
        If Not checker.Match(Kode_Saham.Text).Success OrElse Not checker.Match(Kode_Emiten.Text).Success Then
            Label_ValCode.Text = "Hanya Mengizinkan Angka atau huruf"
            Label_ValCode.Visible = True
            Label_ValNull.Visible = False
            Btn_Donwload.Visible = False
            GridView1.Visible = False
            Kode_Emiten_Src = "xxxxx"
            Kode_Saham_Src = "xxxxx"
        ElseIf Kode_Emiten.Text = "ALL" And Kode_Saham.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = ""
            Kode_Saham_Src = ""
        ElseIf Kode_Emiten.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = ""
            Kode_Saham_Src = Kode_Saham.Text
        ElseIf Kode_Saham.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = Kode_Emiten.Text
            Kode_Saham_Src = ""
        Else
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Saham_Src = Kode_Saham.Text
            Kode_Emiten_Src = Kode_Emiten.Text
            Label_ValNull.Text = ""
        End If
        Dim cmdReader = getsqlquery("SELECT SEC_ISSUER_ID, SEC_ISSUER, SEC_DSC, CODE_BASE_SEC, SEC_MODAL_DASAR, CLO_PRI, TOTAL_BLNC_SCRIP, 
        TOTAL_BLNC_SCRIPLESS FROM CI_SCRIP_SCRIPLESS WHERE SEC_ISSUER_ID Like'%" + Kode_Emiten_Src + "%' AND CODE_BASE_SEC LIKE'%" + Kode_Saham_Src + "%' AND SNAP_DAT LIKE'%" + date_src + "%'", conStringDW)
        GridView1.DataSource = cmdReader
        If cmdReader.Rows.Count = 0 Then
            Label_ValNull.Text = "Data Tidak ditemukan"
            Btn_Donwload.Visible = False
        ElseIf cmdReader.Rows.Count > 0 Then
            Label_ValNull.Text = "Ditemukan : " & cmdReader.Rows.Count & " Data"
            Btn_Donwload.Visible = True
        End If
        GridView1.DataBind()
    End Sub

    <Obsolete>
    Public Sub Generate_Data()
        'QUERY GENERATE ================================================================================================
        Dim Kode_Saham_Src
        Dim Kode_Emiten_Src
        Dim date_src As String
        Dim conv_date
        Dim arr_bulan() As String = {"", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "OCT", "November", "Desember"}
        date_src = ""
        If Snap_Dat.Text <> "" Then
            conv_date = Snap_Dat.Text.Split("-")
            date_src = arr_bulan(conv_date(1)) & "-" & conv_date(0).Substring(2, 2)
        End If
        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"
        Dim checker As Regex = New Regex("^[a-zA-Z0-9]*$", RegexOptions.IgnoreCase)
        If Not checker.Match(Kode_Saham.Text).Success OrElse Not checker.Match(Kode_Emiten.Text).Success Then
            Label_ValCode.Text = "Hanya Mengizinkan Angka atau huruf"
            Label_ValCode.Visible = True
            Label_ValNull.Visible = False
            Btn_Donwload.Visible = False
            GridView1.Visible = False
            Kode_Emiten_Src = "xxxxx"
            Kode_Saham_Src = "xxxxx"
        ElseIf Kode_Emiten.Text = "ALL" And Kode_Saham.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = ""
            Kode_Saham_Src = ""
        ElseIf Kode_Emiten.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = ""
            Kode_Saham_Src = Kode_Saham.Text
        ElseIf Kode_Saham.Text = "ALL" Then
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Emiten_Src = Kode_Emiten.Text
            Kode_Saham_Src = ""
        Else
            Label_ValCode.Visible = False
            Label_ValNull.Visible = True
            GridView1.Visible = True

            Kode_Saham_Src = Kode_Saham.Text
            Kode_Emiten_Src = Kode_Emiten.Text
            Label_ValNull.Text = ""
        End If

        Dim cmdreader = getsqlquery("SELECT sec_issuer_id as kode_emiten, CODE_BASE_SEC as KODE_SAHAM, sec_issuer as nama_emiten, sektor,sub_sektor,sec_modal_dasar as jumlah_saham, clo_pri as closing_price, 
        blnc_lokal_scrip,blnc_asing_scrip,total_blnc_scrip, acct_lokal_scrip, acct_asing_scrip, total_acct_scrip, blnc_lokal_scripless, 
        blnc_asing_scripless, total_blnc_scripless, acct_lokal_scripless, acct_asing_scripless, total_acct_scripless, total_blnc_scrip as perbandingan_scrip,
        total_acct_scripless as perbandingan_scripless, snap_dat FROM CI_SCRIP_SCRIPLESS WHERE SEC_ISSUER_ID Like'%" + Kode_Emiten_Src + "%' AND CODE_BASE_SEC LIKE'%" + Kode_Saham_Src + "%' AND SNAP_DAT LIKE'%" + date_src + "%'", conStringDW)

        'CREATE FILE EXEL================================================================================================
        Try
            Dim Folder_Template As String = "D:\GIT3\Template\"
            Dim Folder_Download As String = "D:\GIT3\Download\"
            Dim Template = Folder_Template & "TEMPLATE_PERBANDINGAN.xlsx"
            Dim Nama_File_Exel = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("dd-MM-yyyy") & "_Exel.xlsx"
            Dim New_Template = Folder_Download & Nama_File_Exel
            Dim File = New FileInfo(Template)
            Dim New_Save_Exel = New FileInfo(New_Template)

            Dim Package_Exel As New ExcelPackage(File)
            Package_Exel.Load(New FileStream(Template, FileMode.Open))

            Dim No As Integer = 1
            Dim Row_Data As Int32 = 9
            ExcelPackage.LicenseContext = LicenseContext.Commercial
            Dim Workshet_Perbandingan As ExcelWorksheet = Package_Exel.Workbook.Worksheets("PERBANDINGAN")
            Dim Moodel_Range = ""
            Dim Model_Table

            For Each data As DataRow In cmdreader.Rows
                Workshet_Perbandingan.Cells("A3").Value = data("SNAP_DAT")
                Workshet_Perbandingan.Cells("A" & Row_Data).Value = data("KODE_EMITEN")
                Workshet_Perbandingan.Cells("B" & Row_Data).Value = data("NAMA_EMITEN")
                Workshet_Perbandingan.Cells("C" & Row_Data).Value = data("JUMLAH_SAHAM")
                Workshet_Perbandingan.Cells("D" & Row_Data).Value = data("CLOSING_PRICE")
                Workshet_Perbandingan.Cells("E" & Row_Data).Value = data("SEKTOR")
                Workshet_Perbandingan.Cells("F" & Row_Data).Value = data("SUB_SEKTOR")

                Workshet_Perbandingan.Cells("G" & Row_Data).Value = data("BLNC_LOKAL_SCRIP")
                Workshet_Perbandingan.Cells("H" & Row_Data).Value = data("BLNC_ASING_SCRIP")
                Workshet_Perbandingan.Cells("I" & Row_Data).Value = data("TOTAL_BLNC_SCRIP")
                Workshet_Perbandingan.Cells("J" & Row_Data).Value = data("ACCT_LOKAL_SCRIP")
                Workshet_Perbandingan.Cells("K" & Row_Data).Value = data("ACCT_ASING_SCRIP")
                Workshet_Perbandingan.Cells("L" & Row_Data).Value = data("TOTAL_ACCT_SCRIP")

                Workshet_Perbandingan.Cells("M" & Row_Data).Value = data("BLNC_LOKAL_SCRIPLESS")
                Workshet_Perbandingan.Cells("N" & Row_Data).Value = data("BLNC_ASING_SCRIPLESS")
                Workshet_Perbandingan.Cells("O" & Row_Data).Value = data("TOTAL_BLNC_SCRIPLESS")
                Workshet_Perbandingan.Cells("P" & Row_Data).Value = data("ACCT_LOKAL_SCRIPLESS")
                Workshet_Perbandingan.Cells("Q" & Row_Data).Value = data("ACCT_ASING_SCRIPLESS")
                Workshet_Perbandingan.Cells("R" & Row_Data).Value = data("TOTAL_ACCT_SCRIPLESS")

                Workshet_Perbandingan.Cells("S" & Row_Data).Value = data("PERBANDINGAN_SCRIP")
                Workshet_Perbandingan.Cells("T" & Row_Data).Value = data("PERBANDINGAN_SCRIPLESS")

                Moodel_Range = "A6:U" & Row_Data
                No = No + 1
                Row_Data = Row_Data + 1
            Next

            Model_Table = Workshet_Perbandingan.Cells(Moodel_Range)
            Model_Table.style.border.top.style = ExcelBorderStyle.Thin
            Model_Table.style.border.left.style = ExcelBorderStyle.Thin
            Model_Table.style.border.right.style = ExcelBorderStyle.Thin
            Model_Table.style.border.bottom.style = ExcelBorderStyle.Thin
            Package_Exel.SaveAs(New_Save_Exel)
        Catch ex As Exception
            Label_Catch.Visible = True
            Label_Catch.Text = ex.Message
        End Try

        'CREATE TAB TXT================================================================================================
        Try
            Dim Name_Txt = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("dd-MM-yyyy")
            Dim Path_Txt As String = "D:\GIT3\Download\" & Name_Txt & "_Tab.txt"
            System.IO.File.Create(Path_Txt).Dispose()

            Dim TargetFile As StreamWriter
            TargetFile = New StreamWriter(Path_Txt, True)
            TargetFile.WriteLine("Data Perbandingan : " & Now())
            TargetFile.WriteLine()
            For Each header As DataColumn In cmdreader.Columns
                TargetFile.Write(header.ColumnName & "|")
            Next
            TargetFile.WriteLine()
            For Each line As DataRow In cmdreader.Rows
                TargetFile.Write(line("KODE_EMITEN") & "|")
                TargetFile.Write(line("NAMA_EMITEN") & "|")
                TargetFile.Write(line("JUMLAH_SAHAM") & "|")
                TargetFile.Write(line("CLOSING_PRICE") & "|")
                TargetFile.Write(line("SEKTOR") & "|")
                TargetFile.Write(line("SUB_SEKTOR") & "|")

                TargetFile.Write(line("BLNC_LOKAL_SCRIP") & "|")
                TargetFile.Write(line("BLNC_ASING_SCRIP") & "|")
                TargetFile.Write(line("TOTAL_BLNC_SCRIP") & "|")
                TargetFile.Write(line("ACCT_LOKAL_SCRIP") & "|")
                TargetFile.Write(line("ACCT_ASING_SCRIP") & "|")
                TargetFile.Write(line("TOTAL_ACCT_SCRIP") & "|")

                TargetFile.Write(line("BLNC_LOKAL_SCRIPLESS") & "|")
                TargetFile.Write(line("BLNC_ASING_SCRIPLESS") & "|")
                TargetFile.Write(line("TOTAL_BLNC_SCRIPLESS") & "|")
                TargetFile.Write(line("ACCT_LOKAL_SCRIPLESS") & "|")
                TargetFile.Write(line("ACCT_ASING_SCRIPLESS") & "|")
                TargetFile.Write(line("TOTAL_ACCT_SCRIPLESS") & "|")

                TargetFile.Write(line("PERBANDINGAN_SCRIP") & "|")
                TargetFile.Write(line("PERBANDINGAN_SCRIPLESS") & "|")

                TargetFile.WriteLine()
            Next
            TargetFile.WriteLine()
            TargetFile.Close()
        Catch ex As Exception
            Label_Catch.Visible = True
            Label_Catch.Text = ex.Message
        End Try

        'CREATE FILE ZIP===============================================================================================
        Dim Zip_Password = "123"
        Try
            'Upload Txt
            Dim Zip_Name_Txt = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("dd-MM-yyyy") & "_Tab.txt"
            Dim Zip_Path_Txt As String = "D:\GIT3\Download\" & Zip_Name_Txt
            'Upload Exel
            Dim Zip_Name_Exel = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("dd-MM-yyyy") & "_Exel.xlsx"
            Dim Zip_Path_Exel As String = "D:\GIT3\Download\" & Zip_Name_Exel

            'Dim Zip_File = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("ddMMyyyyHHmm") & ""
            'Dim Zip_Name As String = [String].Format("" & Zip_File & ".rar")
            'Response.ContentType = "application/7za"
            'Response.AddHeader("content-disposition", "attachment; filename=" + Zip_Name)
            Dim Ex_Zip As New Ionic.Zip.ZipFile()
            Ex_Zip.Password = Zip_Password
            Ex_Zip.AddFile(Zip_Path_Txt)
            Ex_Zip.AddFile(Zip_Path_Exel)
            Ex_Zip.Save("D:\GIT3\Download\LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("ddMMyyyyHHmm") & ".rar")

        Catch ex As Exception
            Label_Catch.Visible = True
            Label_Catch.Text = ex.Message
        End Try

        'CREATE PASSWORD TXT ==========================================================================================
        Dim Zip_Name_Pass = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("dd-MM-yyyy") & "_Pass.txt"
        Dim Zip_Path_Pass As String = "D:\GIT3\Download\" & Zip_Name_Pass
        Try
            System.IO.File.Create(Zip_Path_Pass).Dispose()

            Dim Open_Zip_Pass As StreamWriter
            Open_Zip_Pass = New StreamWriter(Zip_Path_Pass, True)
            Open_Zip_Pass.WriteLine("Data Perbandingan : " & Now())
            Open_Zip_Pass.WriteLine()
            Open_Zip_Pass.WriteLine("Password zip : " & Zip_Password)
            Open_Zip_Pass.Close()

        Catch ex As Exception
            Label_Catch.Visible = True
            Label_Catch.Text = ex.Message
        End Try

        'SEND EMAIL================================================================================================
        Try
            Dim Send_Email As Object
            Send_Email = CreateObject("CDO.Message")
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'or 587
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "gunawanprasetyo313@gmail.com" 'your Google apps mailbox address
            Send_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "gunawan12345" 'Google apps password for that mailbox
            Send_Email.Configuration.Fields.Update

            Send_Email.From = "gunawanprasetyo313@gmail.com"
            Send_Email.To = "reclosher@gmail.com"
            Send_Email.Subject = "Scrip Scrippless Notifikasi"
            Send_Email.AddAttachment(Zip_Path_Pass)
            Send_Email.HTMLBody = "<div>" &
                                "<p>Kepada <b>Unit Penelitian</b>,</p>" &
                                "<div>" &
                                    "<span> Terakait eksekusi procedure untuk data perbandingan script dan scriptless dinyatakan </span>" &
                                    "<b>BERHASIL</b><span>" &
                                    "<span> Dengan Password zip : " & Zip_Password & " </span>" &
                                "</div><br>" &
                                "<div>" &
                                    "<span>Demikian informasi disampaikan,</span><br>" &
                                    "<span>Terima Kasih.</span>" &
                                "</div>" &
                            "</div>"
            Send_Email.Send
        Catch ex As Exception
            Label_Catch.Visible = True
            Label_Catch.Text = ex.Message
        End Try
    End Sub

    Protected Sub ImageButton1_Click1(sender As Object, e As ImageClickEventArgs)
        Dim strPopup As String = "<script language='javascript' ID='ImageButton1'>" + "window.open('Ci_Scrip_Modal.aspx','new window', 'top=90, left=200, width=800, height=400, dependant=no, location=0, alwaysRaised=no, menubar=no, resizeable=no, scrollbars=yes, toolbar=no, status=no, center=yes')" + "</script>"
        ScriptManager.RegisterStartupScript(DirectCast(HttpContext.Current.Handler, Page), GetType(Page), "ImageButton2", strPopup, False)
    End Sub
    Protected Sub ImageButton2_Click1(sender As Object, e As ImageClickEventArgs)
        Dim strPopup As String = "<script language='javascript' ID='ImageButton2'>" + "window.open('Ci_Scrip_Modal.aspx','new window', 'top=90, left=200, width=800, height=400, dependant=no, location=0, alwaysRaised=no, menubar=no, resizeable=no, scrollbars=yes, toolbar=no, status=no, center=yes')" + "</script>"
        ScriptManager.RegisterStartupScript(DirectCast(HttpContext.Current.Handler, Page), GetType(Page), "ImageButton2", strPopup, False)
    End Sub

    <Obsolete>
    Protected Sub Btn_Donwload_Click(sender As Object, e As EventArgs) Handles Btn_Donwload.Click
        Generate_Data()
        Label_ValNull.Visible = False
        Label_ValCode.Visible = False
    End Sub
End Class