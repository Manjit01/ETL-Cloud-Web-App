<%@ Page Language="C#" AutoEventWireup="true" CodeFile="importexporttool.aspx.cs" Inherits="importexporttool" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 100%;
        }

        .auto-style2 {
            text-align: center;
        }

        .auto-style3 {
            font-size: large;
            color: #FF0000;
        }

        .auto-style5 {
            width: 304px;
        }

        .auto-style6 {
            width: 302px;
        }

        .auto-style7 {
            width: 277px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>

            <table class="auto-style1" border="1">
                <tr>
                    <td colspan="4">
                        <h1 class="auto-style2">Tidy ETL Cloud Application</h1>
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6">&nbsp;</td>
                    <td class="auto-style5">&nbsp;</td>
                    <td class="auto-style7">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style6">&nbsp;</td>
                    <td class="auto-style5">&nbsp;</td>
                    <td class="auto-style7">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style6">Select File (*.csv,*.xls,*.xlsx)</td>
                    <td class="auto-style5">&nbsp;</td>
                    <td class="auto-style7">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style6">
                        <asp:FileUpload ID="FileUpload1" runat="server" />
                        <br />
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="FileUpload1" ErrorMessage="Please select a file first" ForeColor="#FF3300" ValidationGroup="group1"></asp:RequiredFieldValidator>
                    </td>
                    <td class="auto-style5">&nbsp;</td>
                    <td class="auto-style7">&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
                <tr>
                    <td class="auto-style6">
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Import Data From Data Source" Height="75px" ValidationGroup="group1" Width="252px" />
                    </td>
                    <td class="auto-style5">
                        <asp:Button ID="btnimporttodb" runat="server" OnClick="btnimporttodb_Click" Text="Save Data To Source Table" Enabled="False" Height="74px" Width="211px" />
                    </td>
                    <td class="auto-style7">
                        <asp:Button ID="btndestindata" runat="server" OnClick="btndestindata_Click" Text="Save Data To Target Table(After Mapping)" Height="70px" Width="248px" />
                    </td>
                    <td>
                        <asp:Button ID="btnExportData" runat="server" OnClick="btnExportData_Click" Text="Export Data" Height="70px" Width="248px" />
                    </td>
                </tr>
                <tr>
                    <td class="auto-style6">
                        <asp:Label ID="lblmsg1" runat="server" CssClass="auto-style3"></asp:Label>
                    </td>
                    <td class="auto-style5">
                        <asp:Label ID="lblmsg2" runat="server" CssClass="auto-style3"></asp:Label>
                    </td>
                    <td class="auto-style7">
                        <asp:Label ID="lblmsg3" runat="server" CssClass="auto-style3"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblmsg4" runat="server" CssClass="auto-style3"></asp:Label>
                        <br />
                        <br />
                                        <asp:HyperLink ID="HyperLink2" runat="server" Target="_blank" Visible="False">Dowload Data File</asp:HyperLink>
                        <br />
                        <asp:Panel ID="Panel1" runat="server" Visible="False">
                            <table class="auto-style1">
                                <tr>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <asp:HyperLink ID="HyperLink1" runat="server" Target="_blank">View Errors</asp:HyperLink>
                                    </td>
                                </tr>
                            </table>
                            Select File (*.csv,*.xls,*.xlsx)<br />
                            <asp:FileUpload ID="FileUpload2" runat="server" />
                            <br />
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="FileUpload2" ErrorMessage="Please select a file first" ForeColor="#FF3300" ValidationGroup="group2"></asp:RequiredFieldValidator>
                            <br />
                            <asp:Button ID="Button2" runat="server" Text="Re-Upload  File" ValidationGroup="group2" OnClick="Button2_Click" />
                            <br />
                        </asp:Panel>
                        <br />
                        <asp:Label ID="lblmsg5" runat="server" CssClass="auto-style3"></asp:Label>
                        <br />
                        <br />
                    </td>
                </tr>
                <tr>
                    <td class="auto-style2" colspan="4">
                        <asp:Image ID="Image1" runat="server" Height="120px" ImageUrl="~/files/giphy.gif" Visible="False" Width="150px" />
                    </td>
                </tr>
            </table>

        </div>
    </form>
</body>
</html>
