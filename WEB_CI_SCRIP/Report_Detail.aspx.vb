Imports System.Data.OracleClient

Public Class Report_Detail
    Inherits System.Web.UI.Page

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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Query_Scrip()
    End Sub

    Sub Query_Scrip()
        Dim id As String
        id = Request.QueryString("Kode_Emiten")
        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"
        Dim data_detail = getsqlquery("Select SEC_ISSUER_ID AS KODE_EMITEN,SEC_ISSUER AS NAMA_EMITEN,CODE_BASE_SEC AS KODE_SAHAM,SEC_DSC AS NAMA_SAHAM, SEKTOR,SUB_SEKTOR,SEC_MODAL_DASAR AS JUMLAH_SAHAM, 
        BLNC_LOKAL_SCRIP,BLNC_ASING_SCRIP,TOTAL_BLNC_SCRIP, ACCT_LOKAL_SCRIP, ACCT_ASING_SCRIP, TOTAL_ACCT_SCRIP, BLNC_LOKAL_SCRIPLESS, 
        BLNC_ASING_SCRIPLESS, TOTAL_BLNC_SCRIPLESS, ACCT_LOKAL_SCRIPLESS, ACCT_ASING_SCRIPLESS, TOTAL_ACCT_SCRIPLESS, TOTAL_BLNC_SCRIP AS PERBANDINGAN_SCRIP,
        TOTAL_ACCT_SCRIPLESS AS PERBANDINGAN_SCRIPLESS FROM CI_SCRIP_SCRIPLESS WHERE SEC_ISSUER_ID='" + id + "'", conStringDW)
        For Each row As DataRow In data_detail.Rows
            Kode_Emiten.Text = row("KODE_EMITEN").ToString()
            Nama_Emiten.Text = row("NAMA_EMITEN").ToString()
            Kode_Saham.Text = row("KODE_SAHAM").ToString()
            Nama_Saham.Text = row("NAMA_SAHAM").ToString()
            Sector.Text = row("SEKTOR").ToString()
            Sub_Sector.Text = row("SUB_SEKTOR").ToString()
            Jumlah_Saham.Text = row("JUMLAH_SAHAM").ToString()

            Scrip_blnc_lokal.Text = row("BLNC_LOKAL_SCRIP").ToString()
            Scrip_blnc_asing.Text = row("BLNC_ASING_SCRIP").ToString()
            Total_blnc_scrip.Text = row("TOTAL_BLNC_SCRIP").ToString()
            Scrip_acct_lokal.Text = row("ACCT_LOKAL_SCRIP").ToString()
            Scrip_acct_asing.Text = row("ACCT_ASING_SCRIP").ToString()
            Total_acct_scrip.Text = row("TOTAL_ACCT_SCRIP").ToString()

            Scripless_blnc_lokal.Text = row("BLNC_LOKAL_SCRIPLESS").ToString()
            Scripless_blnc_asing.Text = row("BLNC_ASING_SCRIPLESS").ToString()
            Total_blnc_scripless.Text = row("TOTAL_BLNC_SCRIPLESS").ToString()
            Scripless_acct_lokal.Text = row("ACCT_LOKAL_SCRIPLESS").ToString()
            Scripless_acct_asing.Text = row("ACCT_ASING_SCRIPLESS").ToString()
            Total_acct_scripless.Text = row("TOTAL_ACCT_SCRIPLESS").ToString()

            Perbandingan_Scrip.Text = row("PERBANDINGAN_SCRIP").ToString()
            Perbandingan_Scripless.Text = row("PERBANDINGAN_SCRIPLESS").ToString()
        Next row
    End Sub

End Class