<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="BranchSeatingLanding.aspx.cs" Inherits="FloorPlan.BranchSeatingLanding" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">

    <script src="Scripts/jquery-ui.min.js" type="text/javascript"></script>
    <script src="Scripts/fixed_table_rc.js" type="text/javascript"></script>
    <script src="Scripts/knockout-min.js" type="text/javascript"></script>
    <link href="Styles/fixed_table_rc.css" rel="stylesheet" type="text/css" />

    <%--    <script src="Scripts/jquery-1.10.0.min.js" type="text/javascript"></script>
    <script src="Scripts/jquery.dataTables.min.js" type="text/javascript"></script>
    <link href="Styles/jquery.dataTables.min.css" rel="stylesheet" type="text/css" />
    <link href="Styles/bootstrap.css" rel="stylesheet" type="text/css" />--%>
    <script type="text/javascript">
        $(document).ready(function () {

            $("#imgHideAll").hide();
        });
        </script>
    <script type="text/javascript">
        function divexpandcollapse(divname) {
            debugger;
            var img = "img" + divname;
            if ($("#" + img).attr("src") == "Image/plus.png") {
                $("#" + img)
				.closest("tr")
				.after("<tr><td></td></tr><tr><td colspan = '100%'>" + $("#" + divname)
				.html() + "</td></tr>");
                $("#" + img).attr("src", "Image/minus.png");
            } else {
                $("#" + img).closest("tr").next().remove();
                $("#" + img).attr("src", "Image/plus.png");
            }
        }
    </script>

    <style type="text/css">
        #grid-container {
            font-family: Calibri;
            font-size: 12px !important;
            line-height: 20px;
        }


        .ft_container table .showImg, .ft_container table .hideImg {
            --padding: 10px;
            cursor: pointer;
        }

        #ft_container table td, #ft_container table th {
            padding: 5px;
        }

        .ft_container table thead {
            background-color: #376091;
            color: White;
        }

        .ft_container table {
            border-collapse: collapse;
            border-spacing: 2px;
            border-color: grey;
            font-size: 12px !important;
           }

        .whiteCell {
            background-color: white;
            width: 1px !important;
        }
    </style>

    <script type="text/javascript">
        function startHide(obj) {
            for (var i = 0; i < obj.length; i++) {
                if (obj[i].ChildRow) {
                    $(".tr_" + obj[i].ID + " .showImg").show();
                    $(".tr_" + obj[i].ID).show();
                }
            }


            for (var i = 0; i < obj.length; i++) {
                if (obj[i].ChildRow) {
                    hideChildren(obj[i]);
                }
            }
        }

        function hideChildren(parentObj) {
            debugger;
            if (parentObj.ChildRow.length > 0) {
                for (var i = 0; i < parentObj.ChildRow.length; i++) {
                    hideChildren(parentObj.ChildRow[i]);
                    //console.log("#tr_"+parentObj.ChildRow[i].ID);
                    $(".tr_" + parentObj.ChildRow[i].ID).hide();
                }
            }
            //console.log("#tr_"+parentObj.ID+" .hideImg");
            $(".tr_" + parentObj.ID + " .hideImg").hide();
            $(".tr_" + parentObj.ID + " .showImg").show();
        }

        function showChildren(parentObj) {
            debugger;
            $(".trParent_" + parentObj.ID).show();
            //            if (parentObj.ChildRow.length > 0) {
            //                for (var i = 0; i < parentObj.ChildRow.length; i++) {
            //                    $(".tr_" + parentObj.ChildRow[i].ID).show();
            //                }
            //            }
            //console.log(".tr_" + parentObj.ID + " .hideImg");
            $(".tr_" + parentObj.ID + " .showImg").hide();
            $(".tr_" + parentObj.ID + " .hideImg").show();
        }

    </script>
    <script type="text/javascript">

        debugger;


        //trim function for ie 7
        // String.prototype.trim = String.prototype.trim || function () { return this.replace(/^\s+|\s+$/g, ''); }

        //var strJson = '[{"ID": 1,"Name": "East","BRANCH_DESIGNED_SEATING": 252,"SEATING_FROM_SPOCS": 210,"DATA_FROM_HR_OPS": 165,"Color": "#6699FF","ParentId": 0, "Level": 0, "ChildRow": [],"ShowChilds": "false" }, {     "ID": 2,  "Name": "North",           "BRANCH_DESIGNED_SEATING": 707,            "SEATING_FROM_SPOCS": 610,            "DATA_FROM_HR_OPS": 385,            "Color": "#6699FF",            "ParentId": 0,            "Level": 0,            "ChildRow": [],            "ShowChilds": "false"        },        {    "ID": 3,    "Name": "South",    "BRANCH_DESIGNED_SEATING": 988,    "SEATING_FROM_SPOCS": 780,    "DATA_FROM_HR_OPS": 695,    "Color": "#6699FF",    "ParentId": 0,    "Level": 0,    "ChildRow": [],    "ShowChilds": "false"}]';
        $(document).ready(function () {
            var strJson = $("#MainContent_hdnJsonStr1").val();

            //   alert(strJson);
            if (strJson.trim() != "") {
                //console.log("Data Present");
                //Knockout Code
                var obj = JSON.parse(strJson);

                function ViewModel() {
                    var self = this;
                    this.Title = "This is Title";
                    this.data = ko.observableArray(obj);
                }
                var viewModel = new ViewModel();


                //Document Ready COde
                $(document).ready(function () {
                    debugger;
                    ko.applyBindings(viewModel);
                    startHide(obj);

                    //console.log("Showing Data");
                    //                $("#gridView").show();

                    // var panelWidth = $("#PanelSBU").width();

                    $('#gridTable').fxdHdrCol({
                        fixedCols: 2,
                        width: "100%",
                        height: 'auto',
                        colModal: [
                           { align: 'left' },
                           { align: 'center' },
                           { align: 'center' }
                        ]
                    });
                });
            }
            else {
                console.log("Data Not Present");
            }

        });

        function ShowAll() {
            debugger;
            for (i = 0; i <= 500; i++) {
               $(".trParent_" + i).show();

                $(".tr_" + i + " .showImg").hide();
                $(".tr_" + i + " .hideImg").show();
            }
            $("#imgShowAll").hide();
            $("#imgHideAll").show();
        }

        function HideAll() {
            debugger;
            for (i = 0; i <= 192; i++) {
                if (parentObj.ChildRow.length > 0) {
                    for (var i = 0; i < parentObj.ChildRow.length; i++) {
                        hideChildren(parentObj.ChildRow[i]);
                        //console.log("#tr_"+parentObj.ChildRow[i].ID);
                        $(".tr_" + parentObj.ChildRow[i].ID).hide();
                    }
                }
                //console.log("#tr_"+parentObj.ID+" .hideImg");
                $(".tr_" + parentObj.ID + " .hideImg").hide();
                $(".tr_" + parentObj.ID + " .showImg").show();
            }
        }
    </script>

    <script type="text/html" id="new-template">
        <tr data-bind="style: { 'color': FontColor, 'background-color': Color }, attr: { 'class': 'tr_grid tr_' + ID + ' trParent_' + ParentId }">
            <td class="tdToggleButton">
                <span data-bind="visible: ChildRow.length > 0">
                    <span class="showImg" style="display: none;" data-bind="attr: { onclick: '$(\'#gridTable .tr_' + ID + ' .showEvent\').click();' }">
                        <img src="Image/plus_small.png" alt="Expand" />
                    </span>
                    <span class="hideImg" style="display: none;" data-bind="attr: { onclick: '$(\'#gridTable .tr_' + ID + ' .hideEvent\').click();' }">
                        <img src="Image/minus_small.png" alt="Collapse" />
                    </span>
                    <span class="showEvent" style="display: none;" data-bind="click: function () { showChildren($data) }"></span>
                    <span class="hideEvent" style="display: none;" data-bind="click: function () { hideChildren($data) }"></span>

                </span>

            </td>
     <%--       <td><span data-bind="text: ID"></span></td>--%>
<%--            <td><span data-bind="text: Name, style: { 'padding-left': (10 * Level) + 'px' }"></span></td>--%>
                        <td> Name: <b data-bind="text: Name"> </b>
        <div data-bind="if:{ Level: 5 }" style="color:red">
          <%--  <u data-bind="text: Level"> </u>--%>
            ABC

          Level  <span data-bind="text: COLOR = '#BDDEFE' ? Level : Name"></span>

        </div>

                       </td>
                            
                            <span data-bind="text: Name, style: { 'padding-left': (10 * Level) + 'px' }"></span></td>
            <td><span data-bind="text: BRANCH_DESIGNED_SEATING"></span></td>
            <td><span data-bind="text: SEATING_FROM_SPOCS"></span></td>
         </tr>

        <!-- ko template: { name: 'new-template', 'foreach': ChildRow } -->
        <!-- /ko -->


       
    </script>

    
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <div id="wrapper">
        <div id="divBranchLanding" runat="server">
            <table>
                <tr>
                  
                    <td>
                        <asp:Label ID="lblRegion" runat="server" Text="Region :" CssClass="DefaultLable"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlRegion" runat="server" CssClass="DefaultDropdown target"
                            AutoPostBack="True" OnTextChanged="ddlRegion_TextChanged">
                        </asp:DropDownList></td>
                    <td>
                        <asp:Label ID="lblState" runat="server" Text="State :" CssClass="DefaultLable"></asp:Label>
                        <asp:HiddenField ID="hdnJsonStr" runat="server" />

                    </td>
                    <td>
                        <asp:DropDownList ID="ddlState" runat="server" CssClass="DefaultDropdown target"
                            AutoPostBack="True" OnTextChanged="ddlState_TextChanged">
                        </asp:DropDownList></td>
                    <td rowspan="2">
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Button" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblCity" runat="server" Text="City :" CssClass="DefaultLable"></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCity" runat="server" CssClass="DefaultDropdown target" AutoPostBack="True" OnSelectedIndexChanged="ddlCity_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td>
                        <asp:Label ID="lblStatus" runat="server" Text="Status :" CssClass="DefaultLable"></asp:Label></td>
                    <td>
                        <asp:DropDownList ID="ddlStatus" runat="server" CssClass="DefaultDropdown target">
                        </asp:DropDownList></td>
                </tr>
            </table>
            <input type="hidden" id="hdnJsonStr1" runat="server">
            <div style="overflow: auto;">
                <asp:GridView ID="gvDisplayData" runat="server" CellSpacing="0" AutoGenerateColumns="False"
                    ShowHeader="false" Width="184px" DataKeyNames="ID">
                    <RowStyle CssClass="RowStyle" Font-Size="11px" />
                    <AlternatingRowStyle CssClass="AlternatingRowStyle" Font-Size="11px" />
                    <Columns>
                        <asp:TemplateField ItemStyle-Width="25px">
                            <ItemTemplate>
                                <div runat="server" id="divGrp">
                                    <a href="JavaScript:divexpandcollapse('div<%# Eval("ID") %>');">
                                        <img alt="Details" id="imgdiv<%# Eval("ID") %>" src="Image/plus.png" style="width: 20px; height: 20px" />
                                    </a>
                                </div>
                                <div id="div<%# Eval("ID") %>" style="display: none;">
                                    <asp:GridView ID="gvchild" runat="server" AutoGenerateColumns="false" DataKeyNames="PARENT_ID"
                                        Width="100%" ShowHeader="false">
                                        <RowStyle CssClass="RowStyle" Font-Size="11px" />
                                        <AlternatingRowStyle CssClass="AlternatingRowStyle" Font-Size="11px" />
                                        <Columns>
                                            <asp:BoundField DataField="NAME" HeaderText="NAME" ItemStyle-Width="70px" />
                                            <asp:BoundField DataField="BRANCH_DESIGNED_SEATING" HeaderText="BRANCH_DESIGNED_SEATING" ItemStyle-Width="70px" />
                                            <asp:BoundField DataField="SEATING_FROM_SPOCS" HeaderText="SEATING_FROM_SPOCS" ItemStyle-Width="70px" />
                                            <asp:BoundField DataField="DATA_FROM_HR_OPS" HeaderText="DATA_FROM_HR_OPS" ItemStyle-Width="70px" />
                                            <asp:BoundField DataField="COLOR" HeaderText="COLOR" ItemStyle-Width="70px" ItemStyle-CssClass="hideGridColumn"
                                                ControlStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn"
                                                FooterStyle-CssClass="hideGridColumn" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="NAME" HeaderText="NAME" ItemStyle-Width="70px" />
                        <asp:BoundField DataField="BRANCH_DESIGNED_SEATING" HeaderText="BRANCH_DESIGNED_SEATING" ItemStyle-Width="70px" />
                        <asp:BoundField DataField="SEATING_FROM_SPOCS" HeaderText="SEATING_FROM_SPOCS" ItemStyle-Width="70px" />
                        <asp:BoundField DataField="DATA_FROM_HR_OPS" HeaderText="DATA_FROM_HR_OPS" ItemStyle-Width="70px" />
                        <asp:BoundField DataField="COLOR" HeaderText="COLOR" ItemStyle-Width="70px" ItemStyle-CssClass="hideGridColumn"
                            ControlStyle-CssClass="hideGridColumn" HeaderStyle-CssClass="hideGridColumn"
                            FooterStyle-CssClass="hideGridColumn" />
                    </Columns>
                </asp:GridView>

                <div id="grid-container" style="width: 500px;height: 600px;overflow: scroll;position: relative;overflow-x: hidden;">
                    <!-- Table to display Grid by KNockout Binding -->
                    <table border="1" id="gridTable" style=" width: 100%;">
                        <thead>
                            <tr>
                                <th>   <img src="Image/plus_small.png" id="imgShowAll" alt="Expand"  onclick="ShowAll();" style="cursor: pointer;"/>
                                    <img src="Image/minus_small.png" id="imgHideAll" alt="Collapse"  onclick="HideAll();" style="cursor: pointer;"/>

                                </th>
                        <%--        <th>ID</th>--%>
                                <th>Name</th>
                                <th>Total Capacity
                                </th>
                                <th>Total Occupancy
                                </th>
                           </tr>
                        </thead>
                        <tbody data-bind="template: { name: 'new-template', 'foreach': data }">
                        </tbody>
                    </table>
                </div>
            </div>

            <asp:PlaceHolder ID="PlaceHolder1" runat="server" />
        </div>
    </div>
</asp:Content>
<%--<script type="text/javascript">
        function ConvertToJqueryTable() {
            $(document).ready(function () {
                alert("Hi1");
                debugger;
                //$('#tblBranchSeating').DataTable();
                //$('#tblBranchSeating1').DataTable();

                var groupColumn = 2;
                var table = $('#tblBranchSeating1').DataTable({
                    "columnDefs": [
                        { "visible": false, "targets": groupColumn }
                    ],
                    "order": [[groupColumn, 'asc']],
                    "displayLength": 25,
                    "drawCallback": function (settings) {
                        var api = this.api();
                        var rows = api.rows({ page: 'current' }).nodes();
                        var last = null;

                        api.column(groupColumn, { page: 'current' }).data().each(function (group, i) {
                            if (last !== group) {
                                $(rows).eq(i).before(
                                    '<tr class="group"><td colspan="5">' + group + '</td></tr>'
                                );

                                last = group;
                            }
                        });
                    }
                });

                // Order by the grouping
                $('#tblBranchSeating1 tbody').on('click', 'tr.group', function () {
                    var currentOrder = table.order()[0];
                    if (currentOrder[0] === groupColumn && currentOrder[1] === 'asc') {
                        table.order([groupColumn, 'desc']).draw();
                    }
                    else {
                        table.order([groupColumn, 'asc']).draw();
                    }
                });

            });
        }
        </script>--%>