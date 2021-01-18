'imports microsoft.visualbasic
'imports system.globalization
'imports system.data
'imports system.web.ui.page
'imports sqlparsing.sqlinjectionparsing
'imports intranet
'imports system.data.odbc

'namespace globaluse
'    public class globalfunction

'        public shared sub showmessagebox(byval message as string, optional byval webpage as system.web.ui.page = nothing)
'            dim lblmessagebox as label = new label()
'            lblmessagebox.text = "<script language='javascript'>" + environment.newline + "window.alert('" + message + "')</script>"
'            if (not webpage is nothing) then
'                webpage.controls.add(lblmessagebox)
'            end if
'        end sub

'        public shared sub setconnectionstring()
'            "persist security info=false;integrated security=sspi;initial catalog=adventureworks;data source=(local)"
'            constringintranet = "password=password;user id=intranet;data source=kseiware.world;persist security info=true"
'            constringintranet = "password=password;user id=intranet;data source=10.111.3.55:1521/kseistpd;persist security info=true"
'            constringintranet = "password=password;user id=intranet;data source=(description = (address = (protocol = tcp)(host = 10.111.3.55)(port = 1521))(connect_data = (server = dedicated) (service_name = kseistpd)));persist security info=true"
'            dim constringintranet = "dsn=intranet;uid=intranet;pwd=password;driver={microsoft odbc for oracle};server=10.111.3.55:1521/kseistpd;"
'            constringintranet = "provider=msdasql.1;user id=intranet;data source=intranet;password=password;"
'            dim constringreportintra = "dsn=intranet;uid=reportintra;pwd=password;driver={microsoft odbc for oracle};server=10.111.3.55:1521/kseistpd;"
'            dim constringintranethukum = "password=password;user id=intranet;data source=(description = (address = (protocol = tcp)(host = 10.111.3.55)(port = 1521))(connect_data = (server = dedicated) (service_name = kseistpd)));persist security info=true"
'            dim constringcaps = "password=caps;user id=caps;data source=intranet.world;persist security info=true"
'            dim constringdw = "password=password;user id=reportintra;data source=kseiware;persist security info=true"
'            dim constringkseiore1 = "password=password;user id=arizona;data source=kseiore1.world;persist security info=true"
'            dim constringkseiore2 = "password=password;user id=oregon;data source=kseiore2.world;persist security info=true"
'            dim constringqueryintranetkseiore1 = "password=password;user id=query_intranet;data source=kseiore1.world;persist security info=true"
'            constringqueryintranetkseiore2 = "password=cbss2000;user id=query_intranet;data source=kseiore2.world;persist security info=true"
'            @20120904 sby : koq ada yang berubah ya dari cbss200 ke queryintranet
'            dim constringqueryintranetkseiore2 = "password=password;user id=query_intranet;data source=kseiore2.world;persist security info=true"
'            dim constringereksa = "password=password;user id=erptreksa;data source=kseiware;persist security info=true"
'            dim strconereksa = "password=password;user id=erptreksa;data source=kseiware;persist security info=true"
'            dim constringintranetoraoledb = "provider=oraoledb.oracle;data source=intranet;user id=intranet;password=password"
'            dim constringstpd = "password=password;user id=onlineksei;data source=kseistpd.world;persist security info=true"
'            dim constringstpi = "password=password;user id=onlineksei;data source=kseistpprod.world;persist security info=true"
'            dim constringfundsep = "password=password;user id=akses;data source=kseistpprod.world;persist security info=true"
'            dim strwebshelldb = "webshell"
'            dim strwebshelluid = "webshell"
'            dim strwebshellpwd = ""
'            dim strdwdb = "password"
'            dim strdwuid = "password"
'            dim strdwpwd = "password"
'            dim constringweb = "password=webksei;user id=webksei;data source=kseiware.world;persist security info=true"
'            dim constringintrans = "password=password;user id=intrans;data source=kseiore1.world;persist security info=true"
'            dim constringdwtrans = "password=password;user id=dwksei;data source=kseiware.world;persist security info=true"

'            --- untuk development ---
'            dim constringdwdev = "password=password;user id=password;data source=intradev.world;persist security info=true"
'            dim constringwebnew = "password=password;user id=webksei;data source=webksei.world;persist security info=true"

'            dim myconstring = "data source=e:\data\public\suitmedia\openweb\app_data\mydatabase1.sdf;password=p@ssw0rd;persist security info=true"

'            dim strsundsn = "sun"
'            strsunuid = "sun"
'            *strsunuid="intra"
'            *strsunpwd="intra"
'            dim strsunuid = "sun"
'            dim strsunpwd = "sunsys"
'            strsunpwd = "rsujoed"
'            strsunpwd = "password"

'            dim strintranetdsn = "intranet"
'            dim strintranetuid = "intranet"
'            dim strintranetpwd = "password"

'            another version of string connection:
'            dim intradb = "intranet"
'            dim intrauserid = "intranet"
'            dim intrauserpwd = "password"

'            ****************** hint for using this file *****************
'            if you find inc\conn.asp included in the script, write codes:
'            con.open intradb, intrauserid, intrauserpwd

'            if you find inc\conndw.asp included in the script, write codes:
'            con.open constringdw

'            ---- digunakan pada aplikasi blokir account (milik bu fitriyah), lihat tabel tt_blokir_acct
'            dim pejabatksei1 = "gusrinaldi akhyar"
'                    dim jabatanksei1 = "kadiv. jasa kustodian sentral"
'                    dim jabatanksei1_en = "head of central depository division"

'                    dim pejabatksei2 = "nina rizalina"
'                    dim jabatanksei2 = "kanit. hubungan pemakai jasa<br>div. jasa kustodian sentral"
'                    dim jabatanksei2_en = "head of customer relation dept.<br>central custodian services division"


'                    dim unitpsdm = "psd"
'                    dim unitkeuangan = "keu"
'        end sub

'#region "connection to db"
'        property providername as string = "oracle.dataaccess.client"
'        private shared odbcconnection as odbcconnection
'        private shared odbccommand as odbccommand

'        private shared intranetlogger = intranetlogservice.instance.getlogger(gettype(globaluse.globalfunction))

'        public shared function sqlparameterparsing(byval parameter as string) as string
'            return sqlinjectionparsing(parameter)
'        end function

'Public Shared Function getsqlquery(ByVal strsql As String, ByVal myconnstring As String) As DataTable
'    Dim result As New DataTable
'    Using connection As New odbcconnection
'        connection.connectionstring = myconnstring
'        connection.open()
'        Using command As New odbccommand(strsql, connection)
'            Using reader As odbcdatareader = command.executereader()
'                result.Load(reader)
'            End Using
'        End Using
'    End Using
'    Return result
'End Function

'        public shared function setsqlquery(byval strsql as string, byval myconnstring as string) as integer
'            dim result as integer
'            using connection as new odbcconnection
'                connection.connectionstring = myconnstring
'                connection.open()
'                using command as new odbccommand(strsql, connection)
'                    result = command.executenonquery()
'                end using
'            end using
'            return result
'        end function
'        for select / get option query
'        public shared function getsql_oradbcast(byval strsql as string, byval myconnstring as string) as odbcdatareader
'            dim odbccommand as odbccommand = getsql_oradbcastcmd(strsql, myconnstring)
'            return odbccommand.executereader()
'        end function

'        for update / delete / insert query
'        public shared function setsql_oradbcast(byval strsql as string, byval myconnstring as string) as integer
'            dim odbccommand as odbccommand = getsql_oradbcastcmd(strsql, myconnstring)
'            return odbccommand.executenonquery()
'        end function

'        public shared function getsql_oradbcastcmd(byval strsql as string, byval myconnstring as string) as odbccommand
'            odbcconnection = new odbcconnection
'            odbcconnection.connectionstring = myconnstring

'            try
'                odbcconnection.open()

'                odbccommand = new odbccommand(strsql, odbcconnection)
'                odbccommand.commandtype = commandtype.text
'                return odbccommand
'            catch ex as odbcexception
'                intranetlogger.error(ex)
'                showmessagebox("get connection: " & ex.message & " > " & ex.source, new page())
'                closeodbcconn()
'                return nothing
'            end try
'        end function

'        for update / delete / insert query
'        public shared function setsql_oradbcastcmd(byval strsql as string, byval myconnstring as string) as odbccommand
'            dim odbccommand as odbccommand = getsql_oradbcastcmd(strsql, myconnstring)
'            return odbccommand
'        end function

'        public shared sub closeodbcconn()
'            if not odbccommand is nothing then odbccommand.dispose()
'            if not odbcconnection is nothing then odbcconnection.close()
'            if not odbcconnection is nothing then odbcconnection.dispose()
'        end sub

'        public shared function getdatenow(optional byval defaultdate as string = "01012001") as string
'            try
'                dim mym as string = month(datetime.now).tostring()
'                dim myd as string = day(datetime.now).tostring()
'                if mym.length = 1 then mym = "0" & mym
'                if myd.length = 1 then myd = "0" & myd
'                return year(datetime.now).tostring() & mym & myd
'            catch ex as odbcexception
'                showmessagebox("cek guest: " & ex.message & " > " & ex.source)
'                return defaultdate
'            end try
'        end function

'        public shared function getoradatenow(optional byval defaultdate as string = "01012001") as string
'            setconnectionstring()
'            dim strquery as string = "select to_char(sysdate,'yyyymmdd') as yyyymmdd from dual"
'            dim myconnstring as string = constringintranet

'            dim dt as datatable = getsqlquery(strquery, myconnstring)
'            for each row as datarow in dt.rows
'                defaultdate = row("yyyymmdd")
'            next row

'            return defaultdate
'        end function
'#end region

'        public shared function filterpassword(byval pw as string, byval pg as page) as boolean
'            dim mynumber as string() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
'            dim mychar as string() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j",
'                                      "k", "l", "m", "n", "o", "p", "q", "r", "s", "t",
'                                      "u", "v", "w", "x", "y", "z"}
'            dim mysymbol as string() = {"~", "!", "@", "#", "$", "%", "^", "&", "*"}

'            dim i as integer
'            dim result as boolean = false
'            dim resulta as boolean = false
'            dim resultb as boolean = false
'            dim resultc as boolean = false

'            for i = 0 to mynumber.length - 1
'                if (pw.indexof(mynumber(i), stringcomparison.ordinalignorecase) >= 0) then
'                    resulta = true
'                    exit for
'                end if
'            next
'            if not resulta then
'                showmessagebox("new password should contains at least one number character!", pg)
'            end if
'            for i = 0 to mychar.length - 1
'                if (pw.indexof(mychar(i), stringcomparison.ordinalignorecase) >= 0) then
'                    resultb = true
'                    exit for
'                end if
'            next
'            if not resultb then
'                showmessagebox("new password should contains at least one alphabets character!", pg)
'            end if
'            for i = 0 to mysymbol.length - 1
'                if (pw.indexof(mysymbol(i), stringcomparison.ordinalignorecase) >= 0) then
'                    resultc = true
'                    exit for
'                end if
'            next
'            if not resultc then
'                showmessagebox("new password should contains at least one symbol character!", pg)
'            end if

'            if resulta and resultb and resultc then
'                result = true
'            end if
'            return result
'        end function
'    end class
'end namespace
