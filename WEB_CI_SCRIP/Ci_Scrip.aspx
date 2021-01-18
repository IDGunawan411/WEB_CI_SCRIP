<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Ci_Scrip.aspx.vb" Inherits="WEB_CI_SCRIP.Ci_Scrip" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   <br><br>
   <table class="">
        <tr>
            <td width="150px"><h5>Kode Emiten</h5></td>
            <td><asp:TextBox ID="Kode_Emiten" runat="server" value="ALL"></asp:TextBox></td>
            <td><asp:ImageButton ID="ImageButton2" OnClick="ImageButton2_Click1" ImageUrl="~/img/search.gif" runat="server" width="20px"/></td>
        </tr>
        <tr>
            <td width="150px"><h5>Kode Saham</h5></td>
            <td><asp:TextBox ID="Kode_Saham" runat="server" value="ALL"></asp:TextBox></td>
            <td><asp:ImageButton ID="ImageButton1" OnClick="ImageButton2_Click1" ImageUrl="~/img/search.gif" runat="server" width="20px"/></td>
        </tr>
        <tr>
            <td width="150px"><h5>Tanggal</h5></td>
            <td><asp:TextBox ID="Snap_Dat" runat="server" type="date" value="ss"></asp:TextBox></td>
        </tr>
        <tr>
            <td width="150px">
            </td>
            <td><asp:Button ID="Button1" runat="server" Height="33px" Text="Search" /></td>
        </tr>
    </table>
    <div class="mt-3 mb-3 container text-center">
        <asp:Label CssClass="text-danger h4" ID="Label_ValNull" runat="server"></asp:Label>
        <asp:Label CssClass="text-danger h4" ID="Label_ValCode" runat="server"></asp:Label>
        <asp:Label CssClass="text-danger h4" ID="Label_Catch" runat="server"></asp:Label>
    </div>
    <div class="table-responsive">
        <asp:GridView ID="GridView1" runat="server" AllowPaging="true" PageSize="10" CssClass="table table-bordered" AutoGenerateColumns="False" align="center" CellPadding="8" ForeColor="#333333">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:TemplateField HeaderText="No">
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Kode Emiten">
                    <ItemTemplate>
                        <a href="Report_Detail.aspx?Kode_Emiten=<%# Eval("SEC_ISSUER_ID") %>"><%# Eval("SEC_ISSUER_ID") %></a>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SEC_ISSUER" HeaderText="Nama Emiten" />

                <asp:BoundField DataField="CODE_BASE_SEC" HeaderText="Kode Saham" />
                <asp:BoundField DataField="SEC_DSC" HeaderText="Nama Saham" />
            
                <asp:BoundField DataField="SEC_MODAL_DASAR" HeaderText="Jumlah Saham" />
                <asp:BoundField DataField="CLO_PRI" HeaderText="Closing Price" />
            
                <asp:BoundField DataField="TOTAL_BLNC_SCRIP" HeaderText="Kepemilikan Saham" DataFormatString="{0:p}"/>
                <asp:BoundField DataField="TOTAL_BLNC_SCRIPLESS" HeaderText="Pemegang Saham" DataFormatString="{0:p}"/>
            
                <%--<asp:BoundField HeaderText="Jumlah" />--%>
            </Columns>
            <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
            <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
            <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
            <SortedAscendingCellStyle BackColor="#FDF5AC" />
            <SortedAscendingHeaderStyle BackColor="#4D0000" />
            <SortedDescendingCellStyle BackColor="#FCF6C0" />
            <SortedDescendingHeaderStyle BackColor="#820000" />
        </asp:GridView>
    </div>
    <div class="text-center"><asp:Button ID="Btn_Donwload" runat="server" Height="33px" Text="Download" /></div>
</asp:Content>
