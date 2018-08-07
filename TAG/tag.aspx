<%@ page language="C#" autoeventwireup="true" inherits="tag, App_Web_nyq0rjcu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <table>
            <tr>
                <td colspan="2"><h4>TAG Maintenance Page</h4></td>
            </tr>
            <tr>
                <td>
                    <asp:ListBox runat="server" ID="tags_lb" Width="175px" Height="200px"  AutoPostBack="true"
                        onselectedindexchanged="tags_lb_SelectedIndexChanged" />   
                        <br /><br />
                        <asp:Button runat="server" ID="addnew_btn" Text="New Tag" 
                        onclick="addnew_btn_Click" /> 
                </td>                            
                <td>
                    <table>
                        <tr>
                            <td>Tag Title</td>
                            <td><asp:TextBox runat="server" ID="tagtitle_tb" Width="200px" /></td>
                        </tr>
                        <tr>
                            <td>Tag Code</td>
                            <td><asp:TextBox runat="server" ID="tagcode_tb" Width="100px" /></td>                            
                        </tr>
                        <tr>
                            <td>Active</td>
                            <td>
                                <asp:RadioButtonList runat="server" ID="tagactive_rbl" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="Yes" Value="1" />
                                    <asp:ListItem Text="No" Value="0" />
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:Button runat="server" ID="submit_btn" Text="Create Tag" />
                                <asp:Button runat="server" ID="update_btn" Text="Update Tag" />
                                <asp:Button runat="server" ID="cancel_btn" Text="Cancel" />
                            </td>
                        </tr>
                    </table>
                </td>                
            </tr>                
        </table>
    </div>
    </form>
</body>
</html>
