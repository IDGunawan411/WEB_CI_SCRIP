Imports System.Data.OracleClient
Imports System.Data

Public Class Ci_Scrip_Modal
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
        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"
        Dim cmdReader = getsqlquery("Select SEC_ISSUER_ID, CODE_BASE_SEC FROM CI_SCRIP_SCRIPLESS", conStringDW)
        GridView1.DataSource = cmdReader
        GridView1.DataBind()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Query_Scrip()
    End Sub
    Sub Query_Scrip()
        Dim Kode_Saham_Src
        Dim Kode_Emiten_Src

        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"

        If Kode_Emiten.Text = "ALL" And Kode_Saham.Text = "ALL" Then
            Kode_Emiten_Src = ""
            Kode_Saham_Src = ""
        ElseIf Kode_Emiten.Text = "ALL" Then
            Kode_Emiten_Src = ""
            Kode_Saham_Src = Kode_Saham.Text
        ElseIf Kode_Saham.Text = "ALL" Then
            Kode_Emiten_Src = Kode_Emiten.Text
            Kode_Saham_Src = ""
        Else
            Kode_Saham_Src = Kode_Saham.Text
            Kode_Emiten_Src = Kode_Emiten.Text
        End If
        Dim cmdReader = getsqlquery("Select SEC_ISSUER_ID, CODE_BASE_SEC FROM CI_SCRIP_SCRIPLESS WHERE SEC_ISSUER_ID Like'%" + Kode_Emiten_Src + "%' AND CODE_BASE_SEC LIKE'%" + Kode_Saham_Src + "%'", conStringDW)

        GridView1.DataSource = cmdReader
        If cmdReader.Rows.Count = 0 Then
            Label1.Text = "Data Tidak ditemukan"
        ElseIf cmdReader.Rows.Count > 0 Then
            Label1.Text = ""
        End If
        GridView1.DataBind()
    End Sub
    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        Query_Scrip()
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)
        Dim pilihan = Request.Form("MyRadioButton")
        Dim conStringDW As String = "Password=password;User ID=reportintra;Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.111.3.55)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME = kseistpd)));"

        Dim strSQL = "Select SEC_ISSUER_ID, CODE_BASE_SEC FROM CI_SCRIP_SCRIPLESS WHERE SEC_ISSUER_ID Like'%" + pilihan + "%'"
        Dim eksekusiSQL = getsqlquery(strSQL, conStringDW)
        Dim Kode_Emiten = eksekusiSQL(0).ItemArray(0)

        Dim Kode_Saham = eksekusiSQL(0).ItemArray(1)

        If Kode_Emiten IsNot Nothing Then
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "script", "opener.document.getElementById('MainContent_Kode_Emiten').value='" & Kode_Emiten & "'", True)
        End If

        If Kode_Saham IsNot Nothing Then
            ClientScript.RegisterClientScriptBlock(Page.GetType(), "script", "opener.document.getElementById('MainContent_Kode_Saham').value='" & Kode_Saham & "'", True)
        End If
        Page.ClientScript.RegisterStartupScript(Page.GetType, "scriptName_popupClose", "<script type=""text/javascript"">close()</script>")
    End Sub
End Class