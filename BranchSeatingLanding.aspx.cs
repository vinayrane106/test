using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Services;
using System.Collections;
using System.Data;
using BusinessLayer;
using common;
using System.Text;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace FloorPlan
{
    public partial class BranchSeatingLanding : System.Web.UI.Page
    {
        BoBranchSeating oBoBranchSeating = new BoBranchSeating();
        clsBranchSeating oclsBranchSeating = new clsBranchSeating();
        #region PageLoad
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                //if (Session["EMP_USER_ID"].ToString() != null)
                //{
                    #region Only for Master
                    System.Web.UI.HtmlControls.HtmlGenericControl HeadLoginView = (System.Web.UI.HtmlControls.HtmlGenericControl)Master.FindControl("HeadLoginViews");
                    HeadLoginView.Visible = false;

                    Label HeadLoginNames = (Label)Master.FindControl("HeadLoginNames");
                    HeadLoginNames.Text = Session["EMP_NAME"].ToString();
                    #endregion

                    if (!IsPostBack)
                    {
                        #region Pageload
                        fillStatusDropDown();
                        fillRegionDropDown();
                        fillBranchSeatingDashboard(Session["EMP_USER_ID"].ToString(), "", "", "");
                        #endregion
                    }
               // }

                //else
                //{
                //    Response.Redirect("~/Login.aspx", true);
                //}

            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        #endregion
        public void fillStatusDropDown()
        {
            try
            {
                DataTable dtStatus = new DataTable();
                string EMP_CODE = Session["EMP_CODE"].ToString();
                dtStatus = oBoBranchSeating.GET_FP_ALL_DROPDOWNS("Status", "", "", EMP_CODE);

                DataSet ds = new DataSet();
                ds.Tables.Add(dtStatus);
                DataView dv = new DataView();
                dv.Table = ds.Tables[0]; // filling dataview with dataset
                dv.Sort = "Status";

                ddlStatus.DataSource = dv;
                ddlStatus.DataTextField = "Status";
                ddlStatus.DataValueField = "Status";
                ddlStatus.DataBind();
                ddlStatus.Items.Insert(0, "--SELECT--");
            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }

        }
        public void fillRegionDropDown()
        {
            try
            {
                ddlState.Items.Clear();
                ddlCity.Items.Clear();

                DataTable dtRegion = new DataTable();
                string EMP_CODE = Session["EMP_CODE"].ToString();
                dtRegion = oBoBranchSeating.GET_FP_ALL_DROPDOWNS("Region", "", "", EMP_CODE);

                DataSet ds = new DataSet();
                ds.Tables.Add(dtRegion);
                DataView dv = new DataView();
                dv.Table = ds.Tables[0]; // filling dataview with dataset
                dv.Sort = "REGION";

                ddlRegion.DataSource = dv;
                ddlRegion.DataTextField = "REGION";
                ddlRegion.DataValueField = "REGION";
                ddlRegion.DataBind();
                ddlRegion.Items.Insert(0, "--SELECT--");
            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        protected void ddlRegion_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ddlState.Items.Clear();
                ddlCity.Items.Clear();

                DataTable dtState = new DataTable();
                string EMP_CODE = Session["EMP_CODE"].ToString();

                if (ddlRegion.SelectedItem.Text != "--SELECT--")
                {
                    dtState = oBoBranchSeating.GET_FP_ALL_DROPDOWNS("State", ddlRegion.SelectedItem.Text, "", EMP_CODE);

                    DataSet ds = new DataSet();
                    ds.Tables.Add(dtState);
                    DataView dv = new DataView();
                    dv.Table = ds.Tables[0]; // filling dataview with dataset
                    dv.Sort = "STATE";

                    ddlState.DataSource = dv;
                    ddlState.DataTextField = "STATE";
                    ddlState.DataValueField = "STATE";
                    ddlState.DataBind();
                    ddlState.Items.Insert(0, "--SELECT--");

                    fillBranchSeatingDashboard(Session["EMP_USER_ID"].ToString(), ddlRegion.SelectedItem.Text, "", "");
                }


            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        protected void ddlState_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //ddlCity.DataSource = null;
                //ddlCity.DataBind();

                ddlCity.Items.Clear();
                DataTable dtCity = new DataTable();
                string EMP_CODE = Session["EMP_CODE"].ToString();

                if (ddlRegion.SelectedItem.Text != "--SELECT--" && ddlState.SelectedItem.Text != "--SELECT--")
                {
                    dtCity = oBoBranchSeating.GET_FP_ALL_DROPDOWNS("City", ddlRegion.SelectedItem.Text, ddlState.SelectedItem.Text, EMP_CODE);

                    DataSet ds = new DataSet();
                    ds.Tables.Add(dtCity);
                    DataView dv = new DataView();
                    dv.Table = ds.Tables[0]; // filling dataview with dataset
                    dv.Sort = "City";

                    ddlCity.DataSource = dv;
                    ddlCity.DataTextField = "CITY";
                    ddlCity.DataValueField = "CITY";
                    ddlCity.DataBind();
                    ddlCity.Items.Insert(0, "--SELECT--");

                    fillBranchSeatingDashboard(Session["EMP_USER_ID"].ToString(), ddlRegion.SelectedItem.Text, ddlState.SelectedItem.Text, "");
                }
            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        public void fillBranchSeatingDashboard(string EMP_CODE, string REGION, string STATE, string CITY)
        {
            try
            {
                DataTable dt = new DataTable();
                dt = oBoBranchSeating.GET_BRANCH_SEATING_DASHBOARD(EMP_CODE, REGION, STATE, CITY);
               // displayGrid(dt);
                var jsonStr = convertDataTableToJson(dt).Replace("'", "");
                hdnJsonStr1.Value = jsonStr;
                Session["JSONDATA"] = jsonStr;
                if (dt.Rows.Count > 0)
                {
                    DataTable parent_dt = dt.Select("LEVEL=0 AND PARENT_ID = 0").CopyToDataTable();
                    ScriptManager.RegisterStartupScript(this, GetType(), "Key", "<script>$('#grid-container').show(); </script>", false);
                }
                else
                {
                    ScriptManager.RegisterStartupScript(this, GetType(), "Key", "<script>$('#grid-container').hide(); </script>", false);
                }
            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        public void displayGrid(DataTable dt)
        {

            //Building an HTML string.
            StringBuilder html = new StringBuilder();

            //Table start.
            html.Append("<table id='tblBranchSeating' border = '1'>");

            //Building the Header row.
            html.Append("<thead>");
            html.Append("<tr>");
          
            foreach (DataColumn column in dt.Columns)
            {
                html.Append("<th>");
                html.Append(column.ColumnName);
                html.Append("</th>");
            }
            html.Append("</thead>");
            html.Append("</tr>");

            //Building the Data rows.
            foreach (DataRow row in dt.Rows)
            {
                html.Append("<tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    html.Append("<td>");
                    html.Append(row[column.ColumnName]);
                    html.Append("</td>");
                }
                html.Append("</tr>");
            }

            //Table end.
            html.Append("</table>");

            //Append the HTML string to Placeholder.
            PlaceHolder1.Controls.Add(new Literal { Text = html.ToString() });
            //ScriptManager.RegisterStartupScript(this, GetType(), "Key", "<script>ConvertToJqueryTable();</script>", false);
        
        }
        private string convertDataTableToJson(DataTable dt)
        {
            string strJson = string.Empty;
            //var returnindex = 0;
            //var parentRows = GetChildRows(dt, 0, out returnindex, 0);
            var parentRows = GetChildRows(dt, 0);
            strJson = JsonConvert.SerializeObject(parentRows, Newtonsoft.Json.Formatting.Indented);
            return strJson;
        }
        private List<ReportViewModel> GetChildRows(DataTable dt, int id)
        {
            List<ReportViewModel> lstReports = new List<ReportViewModel>();
            var childRows = dt.Rows.Cast<DataRow>().Where(r => r["PARENT_ID"].ToString() == id.ToString());
            foreach (var item in childRows)
            {
                var tempReport = fillReportViewModel(item);
                tempReport.ChildRow = GetChildRows(dt, tempReport.ID);
                lstReports.Add(tempReport);
            }
            return lstReports;
        }


        private ReportViewModel fillReportViewModel(DataRow dr)
        {
            ReportViewModel model = new ReportViewModel();
            model.ID = Convert.ToInt32(dr["ID"]);
            model.Name = dr["NAME"].ToString();
            model.BRANCH_DESIGNED_SEATING = Convert.ToInt32(dr["BRANCH_DESIGNED_SEATING"]);
            model.SEATING_FROM_SPOCS = Convert.ToInt32(dr["SEATING_FROM_SPOCS"]);
            model.DATA_FROM_HR_OPS = Convert.ToInt32(dr["DATA_FROM_HR_OPS"]);
            model.Color = dr["COLOR"].ToString();
            model.FontColor = dr["FONTCOLOR"].ToString();
            model.ParentId = Convert.ToInt32(dr["PARENT_ID"]);
            model.Level = Convert.ToInt32(dr["LEVEL"]);
            model.ShowChilds = "false";
            return model;
        }

        public class ReportViewModel
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public int BRANCH_DESIGNED_SEATING { get; set; }
            public int SEATING_FROM_SPOCS { get; set; }
            public int DATA_FROM_HR_OPS { get; set; }
            public string Color { get; set; }
            public string FontColor { get; set; }
            public int ParentId { get; set; }
            public int Level { get; set; }
            public List<ReportViewModel> ChildRow { get; set; }
            public string ShowChilds { get; set; }
        }

        protected void ddlCity_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillBranchSeatingDashboard(Session["EMP_USER_ID"].ToString(), ddlRegion.SelectedItem.Text, ddlState.SelectedItem.Text, ddlCity.SelectedItem.Text);
            }
            catch (Exception ex)
            {
                Response.Redirect("~/Login.aspx", true);
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string Region = "", State = "", City="";
            if (ddlRegion.SelectedItem ==null || ddlRegion.SelectedItem.Text=="--SELECT--")
            {
                Region = ""; 
            }
            if (ddlState.SelectedItem == null || ddlState.SelectedItem.Text == "--SELECT--")
            {
                State = "";
            }
            if (ddlCity.SelectedItem == null || ddlCity.SelectedItem.Text == "--SELECT--")
            {
                City = "";
            }
            dt = oBoBranchSeating.GET_BRANCH_SEATING_DASHBOARD(Session["EMP_USER_ID"].ToString(),Region, State, City);
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            
            Response.Clear();
            Response.Buffer = true;
            //Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", DateTime.Now.ToString("ddMMyyyy")));
            Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", "BranchSeating"));

            //The following directive causes a open/save/cancel dialog for Excel to be displayed
            Response.Cache.SetCacheability(HttpCacheability.Private);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            //Write to response
            Response.BinaryWrite(GetData(ds, "", ""));

            //Response.Flush() 'Do not flush if using compression
            //Response.Close()
            Response.End();
        }


        protected byte[] GetData(DataSet Ds, string LobName, string Version)
        {
            return Export(Ds, LobName, Version);
        }


        public byte[] Export(DataSet DsData, string filter, string Version)
        {
            int i = 1;
            ExcelPackage xlApp = new ExcelPackage();

            foreach (DataTable DtData in DsData.Tables)
            {
                AddSheet(xlApp, DtData, i, filter, Version);
                i++;
            }

            return xlApp.GetAsByteArray();
        }


        protected void AddSheet(ExcelPackage xlApp, DataTable DtData, int Index, string filter, string Version)
        {
            //string TabColor = Convert.ToString(DtData.Rows[0]["TAB_COLOR"]);
            //DtData.Columns.Remove("TAB_COLOR");

            int rowCount = DtData.Rows.Count; int colCount = DtData.Columns.Count;

            string SheetName = "";

            if (Convert.ToString(DtData.TableName).Length > 30)
            {
                SheetName = Convert.ToString(DtData.TableName).Substring(0, 30);
            }
            else
            {
                SheetName = Convert.ToString(DtData.TableName);
            }

            //ExcelWorksheet xlSheet = xlApp.Workbook.Worksheets.Add((DtData.TableName != string.Empty) ? DtData.TableName : string.Format("Sheet{0}", Index.ToString()));
            ExcelWorksheet xlSheet = xlApp.Workbook.Worksheets.Add((SheetName != string.Empty) ? SheetName : string.Format("Sheet{0}", Index.ToString()));
            xlSheet.Cells["A4"].LoadFromDataTable(DtData, true);

            if (filter != "")
                xlSheet.Cells["B2"].Value = filter;
            else
            {
                if (Convert.ToString(DtData.TableName).Substring(0, 1) == "-")
                {
                    xlSheet.Cells["B2"].Value = Convert.ToString(DtData.TableName).Substring(1, Convert.ToString(DtData.TableName).Length - 1) + "-"; // Convert.ToString(DtData.Rows[0]["SBU_LOB_SUBLOB_NAME"]);
                }
                else
                {
                    xlSheet.Cells["B2"].Value = DtData.TableName;
                }
            }

            if (Version != "")
                xlSheet.Cells["B3"].Value = Version;


            //background color
            xlSheet.Cells[2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            xlSheet.Cells[2, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
            xlSheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            xlSheet.Cells[2, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[2, 2].Style.Font.UnderLine = true;
            xlSheet.Cells[2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            xlSheet.Cells[2, 2].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[2, 2].Style.Font.SetFromFont(new Font("Calibri", 9));

            //background color
            xlSheet.Cells[3, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            xlSheet.Cells[3, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
            xlSheet.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            xlSheet.Cells[3, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[3, 2].Style.Font.UnderLine = true;
            xlSheet.Cells[3, 2].Style.Font.SetFromFont(new Font("Calibri", 9));


            //Family = "Calibri";
            //xlSheet.Cells[1, 1, 1, colCount].Merge = true;

            IEnumerable<int> dateColumns = from DataColumn d in DtData.Columns
                                           where d.DataType == typeof(DateTime) || d.ColumnName.Contains("Date")
                                           select d.Ordinal + 1;

            string prevValue = "", currValue = ""; int row = -1;//, colorcolumnNo = 0;
            //  string decimalRow = "";
            var ws = xlSheet;//For Grouping
            ws.OutLineSummaryBelow = false;
            ws.View.FreezePanes(5, 3);
            ws.View.ShowGridLines = false;

            //if (!string.IsNullOrEmpty(TabColor))
            //{
            //    ws.TabColor = System.Drawing.ColorTranslator.FromHtml(TabColor);
            //}

            for (int i = 0; i < rowCount; i++)
            {
                xlSheet.Row(i + 5).Height = 12;
                for (int j = 0; j < colCount; j++)
                {

                    if (Convert.ToString(DtData.Rows[i]["COLOR"]) != "")
                    {
                        xlSheet.Cells[i + 5, j + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        xlSheet.Cells[i + 5, j + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml(Convert.ToString(DtData.Rows[i]["COLOR"])));
                    }

                    //if (DtData.Columns[j].ColumnName == "Rupees in Lakhs")
                    //{
                    //    // merge rows logic
                    //    currValue = xlSheet.Cells[i + 5, j + 1].Value.ToString();
                    //    //     decimalRow = xlSheet.Cells[i + 5, j + 1].Value.ToString();
                    //    if (prevValue == currValue)
                    //    {
                    //        if (row == -1)
                    //            row = i + 4;
                    //    }
                    //    else if (row != -1)
                    //    {
                    //        xlSheet.Cells[row, j + 1, i + 4, j + 1].Merge = true;
                    //        xlSheet.Cells[row, j + 1, i + 4, j + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //        row = -1;
                    //    }

                    //    if (i == rowCount - 1 && prevValue == currValue)//Last Row merging 
                    //    {
                    //        xlSheet.Cells[i + 4, j + 1, i + 5, j + 1].Merge = true;
                    //        xlSheet.Cells[i + 4, j + 1, i + 5, j + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    //    }
                    //    prevValue = currValue;
                    //  }

                    float ft;
                    bool IsNumeric = float.TryParse(Convert.ToString(DtData.Rows[i][j]), out ft);
                    bool IsNumericOtherThanPer = float.TryParse(Convert.ToString(DtData.Rows[i][j]).Replace("%", ""), out ft);

                    //if (!Convert.ToString(DtData.Rows[i][j]).Contains("%") && IsNumeric)//If amount in % then format should be Text or Numeric
                    //{
   
                    //    float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //    if (currValue == "C/I" || currValue == "DE" || currValue == "ROA" || currValue == "ROE (Pre MI)")
                    //    {

                    //        if (currValue == "DE")
                    //        {
                    //            //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //            xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "###,##.0";
                    //            xlSheet.Cells[i + 5, j + 1].Value = realValue;
                    //        }
                    //        else if (currValue == "ROA")
                    //        {
                    //            //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //            realValue = realValue / 100;
                    //            xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "0.0%";
                    //            xlSheet.Cells[i + 5, j + 1].Value = realValue;

                    //        }
                    //        //else if (currValue == "ROE (Pre MI)" || currValue == "C/I")
                    //        //{
                    //        //    //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //        //    realValue = realValue / 100;
                    //        //    xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "0\\%";
                    //        //    xlSheet.Cells[i + 5, j + 1].Value = realValue;

                    //        //}
                    //        else if (currValue == "ROE (Pre MI)")
                    //        {
                    //            //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //            //realValue = realValue * 100;
                    //            realValue = realValue / 100;
                    //            xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "0.0%";
                    //            xlSheet.Cells[i + 5, j + 1].Value = realValue;

                    //        }
                    //        else if (currValue == "C/I")
                    //        {
                    //            //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //            //  realValue = realValue * 100;
                    //            realValue = realValue / 100;
                    //            xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "0%";
                    //            xlSheet.Cells[i + 5, j + 1].Value = realValue;

                    //        }
                    //    }
                    //    else if (currValue == "Avg D/E")
                    //    {
                    //        //  float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]));
                    //        xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "#,##0.0";
                    //        xlSheet.Cells[i + 5, j + 1].Value = realValue;
                    //    }
                    //    else
                    //    {
                    //        xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "###,###0";
                    //        xlSheet.Cells[i + 5, j + 1].Value = realValue;
                    //    }
                    //}

                    //else if (Convert.ToString(DtData.Rows[i][j]).Contains("%") && IsNumericOtherThanPer)
                    //{
                    //    float realValue = float.Parse(Convert.ToString(DtData.Rows[i][j]).Replace("%", ""));
                    //    xlSheet.Cells[i + 5, j + 1].Style.Numberformat.Format = "#0\\.00%";
                    //    xlSheet.Cells[i + 5, j + 1].Value = realValue;


                    //}

                }
                ws.Row(i + 5).OutlineLevel = Convert.ToInt16(DtData.Rows[i]["LEVEL"]);
                ws.Row(i + 5).Collapsed = true;
                xlSheet.Cells[i + 5, 1].Style.Indent = Convert.ToInt16(DtData.Rows[i]["LEVEL"]);

            }


            //(from DataColumn d in DtData.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
            //{
            //    xlSheet.Cells[4, 1, 4, dc].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    xlSheet.Cells[4, 1, 4, dc].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
            //    xlSheet.Cells[4, 1, 4, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    xlSheet.Cells[4, 1, 4, dc].Style.Font.Color.SetColor(System.Drawing.Color.White);

            //    //border
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Top.Color.SetColor(System.Drawing.Color.White);
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
            //    xlSheet.Cells[4, dc, rowCount + 4, dc].Style.Border.Left.Color.SetColor(System.Drawing.Color.White);
            //});


            xlSheet.Cells[4, 1, 4, colCount].Style.Fill.PatternType = ExcelFillStyle.Solid;
            xlSheet.Cells[4, 1, 4, colCount].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
            xlSheet.Cells[4, 1, 4, colCount].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            xlSheet.Cells[4, 1, 4, colCount].Style.Font.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            //border
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Top.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Border.Left.Color.SetColor(System.Drawing.Color.White);
            //xlSheet.Cells[4, 1, rowCount + 4, colCount].AutoFitColumns();
            xlSheet.Cells[4, 1, rowCount + 4, colCount].Style.Font.SetFromFont(new Font("Calibri", 9));
            xlSheet.Cells[5, 3, rowCount + 4, colCount].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            xlSheet.Cells[5, 1, rowCount + 4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            //xlSheet.Cells[5, 1, rowCount + 4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //xlSheet.Cells[5, 1, rowCount + 4, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);







            for (int i = 1; i <= colCount; i++)
                xlSheet.Column(i).Width = 10;
            for (int i = 1; i <= 5; i++)
                xlSheet.Row(4).Height = 12;

            xlSheet.Column(2).Width = 30;
            // hide color column
            //string[] hideColumnNames = { "ID", "PARENT_ID", "COLOR", "LEVEL" };
            //IEnumerable<int> hideColumnNos = from DataColumn d in DtData.Columns
            //                                 where hideColumnNames.Contains(d.ColumnName)
            //                                 select d.Ordinal + 1;
            //foreach (int dc in hideColumnNos)
            //{
            //    xlSheet.Column(dc).Hidden = true;

            //}


            // This section is varies during May-June Month :" Start ": Done by Priyadarshan 2-May-2016
            //  string[] deleteColumnNames = { "Budget FTY", "Budget Q1", "Budget Q2", "Budget Q3", "Budget Q4", "ID", "PARENT_ID", "COLOR", "LEVEL" };
            string[] deleteColumnNames = { "ID", "PARENT_ID", "COLOR", "LEVEL", "FONTCOLOR" };
            IEnumerable<int> deleteColumnNos = from DataColumn d in DtData.Columns
                                               where deleteColumnNames.Contains(d.ColumnName)
                                               select d.Ordinal + 1;
            foreach (int dc in deleteColumnNos)
            {
                ws.DeleteColumn(dc);


            }
            // This section is varies during May-June Month :" End ": Done by Priyadarshan 2-May-2016


            //string[] deleteColumnNames = { "ID", "PARENT_ID" ,"COLOR", "LEVEL" };
            //IEnumerable<int> deleteColumnNos = from DataColumn d in DtData.Columns
            //                                 where deleteColumnNames.Contains(d.ColumnName)
            //                                 select d.Ordinal + 1;
            //foreach (int dc in deleteColumnNos)
            //{
            //    ws.DeleteColumn(dc);

            //}

            //Blank Column.
            ws.Column(16).Width = 1;
            ws.Column(16).Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Column(16).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            //ws.Column(20).Width = 1;
            //ws.Column(20).Style.Fill.PatternType = ExcelFillStyle.Solid;
            //ws.Column(20).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

            if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells["Q4"].Value)))
            {
                ws.Cells["Q3:R3"].Merge = true;
                ws.Cells["Q3:R3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["Q3:R3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
                ws.Cells["Q3:R3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["Q3:R3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["Q3:R3"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["Q3:R3"].Style.Font.UnderLine = true;
                ws.Cells["Q3:R3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["Q3:R3"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["Q3:R3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["Q3:R3"].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["Q3:R3"].Style.Font.SetFromFont(new Font("Calibri", 9));
                ws.Cells["Q3:R3"].Value = "Q1";
            }

            if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells["S4"].Value)))
            {
                ws.Cells["S3:T3"].Merge = true;
                ws.Cells["S3:T3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["S3:T3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
                ws.Cells["S3:T3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["S3:T3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["S3:T3"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["S3:T3"].Style.Font.UnderLine = true;
                ws.Cells["S3:T3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["S3:T3"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["S3:T3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["S3:T3"].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["S3:T3"].Style.Font.SetFromFont(new Font("Calibri", 9));
                ws.Cells["S3:T3"].Value = "Q2";
            }

            if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells["U4"].Value)))
            {
                ws.Cells["U3:V3"].Merge = true;
                ws.Cells["U3:V3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["U3:V3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
                ws.Cells["U3:V3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["U3:V3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["U3:V3"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["U3:V3"].Style.Font.UnderLine = true;
                ws.Cells["U3:V3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["U3:V3"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["U3:V3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["U3:V3"].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["U3:V3"].Style.Font.SetFromFont(new Font("Calibri", 9));
                ws.Cells["U3:V3"].Value = "Q3";
            }

            if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells["W4"].Value)))
            {
                ws.Cells["W3:X3"].Merge = true;
                ws.Cells["W3:X3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["W3:X3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#376091"));
                ws.Cells["W3:X3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["W3:X3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["W3:X3"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["W3:X3"].Style.Font.UnderLine = true;
                ws.Cells["W3:X3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["W3:X3"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["W3:X3"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["W3:X3"].Style.Border.Right.Color.SetColor(System.Drawing.Color.White);
                ws.Cells["W3:X3"].Style.Font.SetFromFont(new Font("Calibri", 9));
                ws.Cells["W3:X3"].Value = "Q4";
            }


        }



    }
}