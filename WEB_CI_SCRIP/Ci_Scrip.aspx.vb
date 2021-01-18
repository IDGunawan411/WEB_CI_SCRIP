Imports System.Data.OracleClient
Imports System.Data
Imports OfficeOpenXml
Imports System.IO
Imports OfficeOpenXml.Style

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
    Sub Generate_Exel()

        Dim foldertemplate As String = "D:\GIT3\Template\"
        Dim folderdownload As String = "D:\GIT3\Download\"
        Dim template1 = foldertemplate & "TEMPLATE_PERBANDINGAN.xlsx"
        Dim newfile = "LAPORAN_PERBANDINGAN_" & DateTime.Now.ToString("ddMMyyyyHHmmss") & ""
        newfile = newfile & ".xlsx"
        Dim newtemplate = folderdownload & newfile
        Dim File = New FileInfo(template1)
        Dim newfiles = New FileInfo(newtemplate)
        Try
            Using package As New ExcelPackage(File)
                package.Load(New FileStream(template1, FileMode.Open))

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

                Dim no As Integer = 1
                Dim row_data As Int32 = 9
                ExcelPackage.LicenseContext = LicenseContext.Commercial
                Dim wk1 As ExcelWorksheet = package.Workbook.Worksheets("PERBANDINGAN")
                Dim modelrange = ""
                Dim modeltable

                For Each data As DataRow In cmdreader.Rows
                    wk1.Cells("A3").Value = data("SNAP_DAT")
                    wk1.Cells("A" & row_data).Value = data("KODE_EMITEN")
                    wk1.Cells("B" & row_data).Value = data("NAMA_EMITEN")
                    wk1.Cells("C" & row_data).Value = data("JUMLAH_SAHAM")
                    wk1.Cells("D" & row_data).Value = data("CLOSING_PRICE")
                    wk1.Cells("E" & row_data).Value = data("SEKTOR")
                    wk1.Cells("F" & row_data).Value = data("SUB_SEKTOR")

                    wk1.Cells("G" & row_data).Value = data("BLNC_LOKAL_SCRIP")
                    wk1.Cells("H" & row_data).Value = data("BLNC_ASING_SCRIP")
                    wk1.Cells("I" & row_data).Value = data("TOTAL_BLNC_SCRIP")
                    wk1.Cells("J" & row_data).Value = data("ACCT_LOKAL_SCRIP")
                    wk1.Cells("K" & row_data).Value = data("ACCT_ASING_SCRIP")
                    wk1.Cells("L" & row_data).Value = data("TOTAL_ACCT_SCRIP")

                    wk1.Cells("M" & row_data).Value = data("BLNC_LOKAL_SCRIPLESS")
                    wk1.Cells("N" & row_data).Value = data("BLNC_ASING_SCRIPLESS")
                    wk1.Cells("O" & row_data).Value = data("TOTAL_BLNC_SCRIPLESS")
                    wk1.Cells("P" & row_data).Value = data("ACCT_LOKAL_SCRIPLESS")
                    wk1.Cells("Q" & row_data).Value = data("ACCT_ASING_SCRIPLESS")
                    wk1.Cells("R" & row_data).Value = data("TOTAL_ACCT_SCRIPLESS")

                    wk1.Cells("S" & row_data).Value = data("PERBANDINGAN_SCRIP")
                    wk1.Cells("T" & row_data).Value = data("PERBANDINGAN_SCRIPLESS")

                    modelrange = "A6:U" & row_data
                    no = no + 1
                    row_data = row_data + 1
                Next

                modeltable = wk1.Cells(modelrange)
                'assign(borders)
                modeltable.style.border.top.style = ExcelBorderStyle.Thin
                modeltable.style.border.left.style = ExcelBorderStyle.Thin
                modeltable.style.border.right.style = ExcelBorderStyle.Thin
                modeltable.style.border.bottom.style = ExcelBorderStyle.Thin
                package.SaveAs(newfiles)
            End Using
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
        Generate_Exel()
        Label_ValNull.Visible = False
        Label_ValCode.Visible = False
    End Sub
End Class