<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="pptx2ppt._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <h1>PPT to PPTX converter</h1>
    <asp:Panel ID="pnlInteractive" runat="server" Visible="False">
        <p>Use local (at server) path to source and destination files using double backslashes (\\); exampli gratia in the textboxes bellow.<br /></p>
        <input id="txbInput" name="input" placeholder="C:\\Windows\\Temp\\convert-this.pptx" style="width:100%;" onchange="javascript:translateURI();" />
        <input id="txbOutput" name="output" placeholder="C:\\Windows\\Temp\\converted.pptx" style="width:100%;" onchange="javascript:translateURI();" />
        <div style="text-align:right;">
            <input type="submit" value="Convert" />
        </div>
        <script type="text/javascript">
            self.translateURI = function () {
                var compiledURI = location.protocol + '//' + location.host + location.pathname + '?input=' + document.getElementById('txbInput').value + '&output=' + document.getElementById('txbOutput').value;
                var linkToURI = '<a href="'+compiledURI+'">'+compiledURI+'</a>';
                document.getElementById("MainContent_lblDebug").innerHTML = 'URI to this convertion:' + '\n' + linkToURI;
            }
        </script>
    </asp:Panel>
    <asp:Label ID="lblDebug" runat="server" Text="" Style="display:block; margin:1ex 0; padding:0.25em; font-family:monospace; white-space:pre; border:1px solid #808080; background-color:#FFFFCC; overflow:visible;"></asp:Label>
</asp:Content>
