Imports System.Data.OracleClient

Namespace GlobalUse
    Public Class GlobalConnection

        Public Shared Sub ShowMessageBox(ByVal Message As String, Optional ByVal WebPage As System.Web.UI.Page = Nothing)
            Dim lblMessageBox As Label = New Label()

            lblMessageBox.Text =
                "<script language='javascript'>" + Environment.NewLine +
                "window.alert('" + Message + "')</script>"
            If (Not WebPage Is Nothing) Then
                WebPage.Controls.Add(lblMessageBox)
            End If
        End Sub

        Public Shared Sub SetConnectionString()
            '"Persist Security Info=False;Integrated Security=SSPI;Initial Catalog=AdventureWorks;Data Source=(local)"
            Dim conStringIntranet = "Provider=MSDAORA.1;Password=password;User ID=intranet;Data Source=KSEISTPD;Persist Security Info=True"
            Dim conStringCaps = "Password=caps;User ID=caps;Data Source=intranet.world;Persist Security Info=True"
            Dim conStringDW = "Provider=MSDAORA.1;Password=password;User ID=reportintra;Data Source=KSEISTPD;Persist Security Info=True"
            'conStringDW = "Password=password;User ID=reportintra;Data Source=kseiware.world;Persist Security Info=True"
            Dim conStringKseiore1 = "Password=password;User ID=arizona;Data Source=kseiore1.world;Persist Security Info=True"
            Dim conStringKseiore2 = "Password=password;User ID=oregon;Data Source=kseiore2.world;Persist Security Info=True"
            Dim conStringQueryIntranetKseiore1 = "Password=password;User ID=query_intranet;Data Source=kseiore1.world;Persist Security Info=True"
            'conStringQueryIntranetKseiore2="Password=cbss2000;User ID=query_intranet;Data Source=kseiore2.world;Persist Security Info=True"
            '@20120904 SBY : koq ada yang berubah ya dari cbss200 ke queryintranet
            Dim conStringQueryIntranetKseiore2 = "Password=password;User ID=query_intranet;Data Source=kseiore2.world;Persist Security Info=True"
            Dim conStringEReksa = "Password=password;User ID=erptreksa;Data Source=kseiware.world;Persist Security Info=True"
            Dim strConEReksa = "Password=password;User ID=erptreksa;Data Source=kseiware.world;Persist Security Info=True"
            Dim conStringIntranetOraOleDB = "Provider=OraOLEDB.Oracle;Data Source=intranet.WORLD;User ID=intranet;PASSWORD=password"
            Dim conStringSTPD = "Password=password;User ID=onlineksei;Data Source=kseistpd.world;Persist Security Info=True"
            Dim conStringSTPI = "Password=password;User ID=onlineksei;Data Source=kseistpprod.world;Persist Security Info=True"
            Dim conStringFundSep = "Password=password;User ID=akses;Data Source=kseistpprod.world;Persist Security Info=True"
            Dim strWebshellDB = "webshell"
            Dim strWebshellUID = "webshell"
            Dim strWebshellPWD = ""
            Dim strDWDB = "password"
            Dim strDWUID = "password"
            Dim strDWPWD = "password"
            Dim conStringWeb = "Password=webksei;User ID=webksei;Data Source=kseiware.world;Persist Security Info=True"
            Dim conStringIntrans = "Password=password;User ID=intrans;Data Source=kseiore1.world;Persist Security Info=True"
            Dim conStringDWTrans = "Password=password;User ID=dwksei;Data Source=kseiware.world;Persist Security Info=True"

            '--- untuk development ---
            Dim conStringDWDev = "Password=password;User ID=password;Data Source=intradev.world;Persist Security Info=True"
            Dim conStringWebNew = "Password=password;User ID=webksei;Data Source=webksei.world;Persist Security Info=True"

            Dim myConString = "Data Source=E:\Data\Public\Suitmedia\OpenWeb\App_Data\MyDatabase1.sdf;Password=p@ssw0rd;Persist Security Info=True"

            Dim strSunDSN = "sun"
            'strSunUID="sun"
            '*strSunUID="intra"
            '*strSunPWD="intra"
            Dim strSunUID = "sun"
            Dim strSunPWD = "sunsys"
            'strSunPWD="rsujoed"
            'strSunPWD="password"

            Dim strIntranetDSN = "intranet"
            Dim strIntranetUID = "intranet"
            Dim strIntranetPWD = "password"

            'Another version of string connection:
            'dim intraDB = "intranet"
            'dim intraUserID = "intranet"
            'dim intraUserPwd = "password"

            '****************** HINT for using this file *****************
            'If You find inc\conn.asp included in the script, write codes:
            'con.open intraDB,intraUserID,intraUserPwd

            'if You find inc\connDW.asp included in the script, write codes:
            'con.open conStringDW

            '---- Digunakan pada aplikasi Blokir Account (milik Bu Fitriyah), lihat tabel TT_BLOKIR_ACCT
            Dim pejabatKSEI1 = "Gusrinaldi Akhyar"
            Dim JabatanKSEI1 = "Kadiv. Jasa Kustodian Sentral"
            Dim JabatanKSEI1_en = "Head of Central Depository Division"

            Dim pejabatKSEI2 = "Nina Rizalina"
            Dim JabatanKSEI2 = "Kanit. Hubungan Pemakai Jasa<br>Div. Jasa Kustodian Sentral"
            Dim JabatanKSEI2_en = "Head of Customer Relation Dept.<br>Central Custodian Services Division"


            Dim unitPSDM = "PSD"
            Dim unitKeuangan = "KEU"
        End Sub

        'For Select / Get Option Query
        Public Shared Function GetSQL_OraDBCast(ByVal strSQL As String, ByVal myconnString As String) As OracleDataReader
            Dim OraConnection As New OracleConnection
            OraConnection.ConnectionString = myconnString

            Try
                OraConnection.Open()

                Dim OraCommand As New OracleCommand(strSQL, OraConnection)
                OraCommand.CommandType = CommandType.Text

                Return OraCommand.ExecuteReader()
            Catch ex As OracleException
                ShowMessageBox("Get Connection: " & ex.Message & " > " & ex.Source)
                Return Nothing
            End Try

            OraConnection.Close()
            OraConnection.Dispose()
        End Function

        'For Update/Delete/Insert Query
        Public Shared Function SetSQL_OraDBCast(ByVal strSQL As String, ByVal myconnString As String) As Integer
            Dim OraConnection As New OracleConnection
            OraConnection.ConnectionString = myconnString

            Try
                OraConnection.Open()

                Dim OraCommand As New OracleCommand(strSQL, OraConnection)
                OraCommand.CommandType = CommandType.Text

                Return OraCommand.ExecuteNonQuery()
            Catch ex As OracleException
                ShowMessageBox("Set Connection: " & ex.Message & " > " & ex.Source)
                Return Nothing
            End Try

            OraConnection.Close()
            OraConnection.Dispose()
        End Function

    End Class
End Namespace
