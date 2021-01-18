<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Ci_Scrip_Modal.aspx.vb" Inherits="WEB_CI_SCRIP.Ci_Scrip_Modal" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <table class="">
        <tr>
            <td width="150px"><h5>Kode Emiten</h5></td>
            <td><asp:TextBox ID="Kode_Emiten" runat="server" value="ALL"></asp:TextBox></td>
        </tr>
        <tr>
            <td width="150px"><h5>Kode Saham</h5></td>
            <td><asp:TextBox ID="Kode_Saham" runat="server" value="ALL"></asp:TextBox></td>
        </tr>
        <tr>
            <td width="150px">
            <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Height="33px" Text="Select" />
            </td>
            <td><asp:Button ID="Button1" runat="server" Height="33px" Text="Search" /></td>
        </tr>
    </table>
        <div class="mt-3 mb-3 container text-center">
        <asp:Label CssClass="text-danger h4" ID="Label1" runat="server"></asp:Label>
    </div>
    <div class="table-responsive">
        <asp:GridView ID="GridView1" runat="server" AllowPaging="true" PageSize="10" CssClass="table table-bordered" AutoGenerateColumns="False" align="center" CellPadding="8" ForeColor="#333333">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:TemplateField HeaderText="Select One">
                    <ItemTemplate >
                        <input name="MyRadioButton" type="radio"
                            value='<%# Eval("SEC_ISSUER_ID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="SEC_ISSUER_ID" HeaderText="Kode Emiten" />
                <asp:BoundField DataField="CODE_BASE_SEC" HeaderText="Kode Saham" />
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
</asp:Content>