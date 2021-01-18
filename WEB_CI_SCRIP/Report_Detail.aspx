<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Report_Detail.aspx.vb" Inherits="WEB_CI_SCRIP.Report_Detail" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <style>
        .sub-content {
            border: 1px solid gray;
            box-shadow :inherit ;
            padding : 5px;
        }
        .content{
            border: 1px solid gray;
            padding : 50px;
            margin-top : 50px;
        }
        .auto-style1 {
            width: 717px;
        }
        .auto-style2 {
            width: 304px;
        }
        .auto-style3 {
            width: 281px;
        }
    </style>
    <div class="content"> 
        <h4><b>Report Detail</b></h4><br>
        <table class="">
            <tr>
                <td width="30%"><h5><b>Kode Emiten</b></h5></td>
                <td width="10px">:</td>
                <td class="auto-style1"><asp:Label ID="Kode_Emiten" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td><h5><b>Nama Emiten</b></h5></td>
                <td>:</td>
                <td class="auto-style1"><asp:Label ID="Nama_Emiten" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td><h5><b>Kode Saham</b></h5></td>
                <td>:</td>
                <td class="auto-style1"><asp:Label ID="Kode_Saham" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td><h5><b>Nama Saham</b></h5></td>
                <td>:</td>
                <td class="auto-style1"><asp:Label ID="Nama_Saham" runat="server"></asp:Label></td>
            </tr>
            <tr><td></td><td></td></tr>
            <tr>
                <td><h5><b>Sector</b></h5></td>
                <td>:</td>
                <td class="auto-style1"><asp:Label ID="Sector" runat="server">aa</asp:Label></td>
            </tr>
            <tr>
                <td><h5><b>Sub Sector</b></h5></td>
                <td>:</td>
                <td class="auto-style1"><asp:Label ID="Sub_Sector" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td><h5><b>Jumlah Saham Beredar</b></h5></td>
                <td>:</td>
                <td align="right" class="auto-style1"><asp:Label ID="Jumlah_Saham" runat="server"></asp:Label></td>
            </tr>
        </table>
        <br>
        <%--Scrip--%>
        <span><b>Script</b></span>
        <div class="sub-content">
            <table class="">
                <tr>
                    <td class="auto-style2"><h5><b>Kepemilikan : </b></h5></td>
                    <td width="2%"></td><td class="auto-style1"></td>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Lokal</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scrip_blnc_lokal" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scrip_blnc_asing" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Total_blnc_scrip" runat="server"></asp:Label></td>
                </tr>

                <tr>
                    <td class="auto-style2"><h5><b>Pemegang Saham : </b></h5></td>
                    <td></td><td></td>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Lokal</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scrip_acct_lokal" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scrip_acct_asing" runat="server"></asp:Label></t>
                </tr>
                <tr>
                    <td class="auto-style2"><h5>Total</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Total_acct_scrip" runat="server"></asp:Label></td>
                </tr>
            </table>
        </div>
        <br>
        <%--Scripless--%>
        <span><b>Scriptless</b></span>
        <div class="sub-content">
            <table class="">
                <tr>
                    <td class="auto-style2"><h5><b>Kepemilikan : </b></h5></td>
                    <td></td><td class="auto-style1"></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Lokal</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scripless_blnc_lokal" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scripless_blnc_asing" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Total_blnc_scripless" runat="server"></asp:Label></td>
                </tr>

                <tr>
                    <td class="auto-style3"><h5><b>Pemegang Saham : </b></h5></td>
                    <td></td><td></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Lokal</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scripless_acct_lokal" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Asing</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Scripless_acct_asing" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="auto-style3"><h5>Total</h5></td>
                    <td width="10px">:</td>
                    <td align="right"><asp:Label ID="Total_acct_scripless" runat="server"></asp:Label></td>
                </tr>
            </table>
        </div>
        <br>
        <%--Scrip Scripless--%>
        <span><b>Scrip & Scriptless</b></span>
        <div class="sub-content">
            <table class="">
                <tr>
                    <td width="500px"><h5><b>Hasil Perbandingan Scrip</b></h5></td>
                    <td></td>
                    <td align="right" class="auto-style1"><asp:Label ID="Perbandingan_Scrip" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td width=""><h5><b>Hasil Perbandingan Scripless</b></h5></td>
                    <td width="10px"></td>
                    <td align="right"><asp:Label ID="Perbandingan_Scripless" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td width=""><h5><b>Selisih Julmah saham yang beredar dengan total saham di Scrip & Scripless</b></h5></td>
                    <td width="10px"></td>
                    <td align="right"><asp:Label ID="Selisih" runat="server"></asp:Label></t>
                </tr>
            </table>
        </div>
        <br>
        
        <a href="Ci_Scrip.aspx" class="btn btn-danger btn-sm">BACK</a>
   </div>
</asp:Content>
