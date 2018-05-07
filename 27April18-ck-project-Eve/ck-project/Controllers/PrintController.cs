
using ck_project.Helpers;
using ck_project.Models;
using PagedList;
using SelectPdf;
using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Dynamic;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

namespace ck_project.Controllers
{
    [Authorize]
    public class PrintController : Controller
    {
        ckdatabase db = new ckdatabase();
        ProjSummaryHelper projSummaryHelper = new ProjSummaryHelper();

        public ActionResult Convert(string documentName, int id, string str)
        {
            // get the data
            var projSummary = new ProjectSummary
            {
                Branch = db.branches.ToList(),
            };
            var lead = db.leads.Where(l => l.lead_number == id).FirstOrDefault();

            if (lead != null)
            {
                projSummary.Lead = lead;
                projSummary = projSummaryHelper.SetCustomerData(lead, projSummary);
                projSummary = projSummaryHelper.SetAddresses(lead, projSummary);
                projSummary = projSummaryHelper.GetProductCategoryList(lead, projSummary);
                projSummary = projSummaryHelper.GetProductTotalMap(lead, projSummary);
                projSummary = projSummaryHelper.CalculateProposalAmtDue(id, projSummary);
                projSummary = projSummaryHelper.CalculateInstallCategoryCostMap(lead, projSummary);
                projSummary = projSummaryHelper.CalculateInstallationsData(lead, projSummary);
            }

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView(str, projSummary);

            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert1(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM lead_source_report";

            var model = new List<CLSA>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA();
                        obj.source_name = rdr["source_name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("CLSAPrint", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert1Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM lead_source_report";

            var model = new List<CLSA>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA();
                        obj.source_name = rdr["source_name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];
                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        public ActionResult Convert2(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA2>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA2();
                        obj.ProjectName = rdr["Project Type Name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];
                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View2Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert2Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA2>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA2();
                        obj.ProjectName = rdr["Project Type Name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];
                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        public ActionResult Convert3(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA3>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA3();
                        obj.Project_Status_Name = rdr["Project Status Name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];
                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View3Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert3Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA3>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA3();
                        obj.Project_Status_Name = rdr["Project Status Name"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.total = (int)rdr["total"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        public ActionResult Convert4(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA4>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA4();
                        obj.Type = rdr["Types"].ToString();
                        obj.Huntington = (double)rdr["Huntington"];
                        obj.Charleston = (double)rdr["Charleston"];
                        obj.Lewisburg = (double)rdr["Lewisburg"];
                        obj.Companytotal = (double)rdr["Company Total"];
                        obj.Percentage = (double)rdr["Percentage"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View4Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert4Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA4>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {

                        var obj = new CLSA4();
                        obj.Type = rdr["Types"].ToString();
                        obj.Huntington = (double)rdr["Huntington"];
                        obj.Charleston = (double)rdr["Charleston"];
                        obj.Lewisburg = (double)rdr["Lewisburg"];
                        obj.Companytotal = (double)rdr["Company Total"];
                        obj.Percentage = (double)rdr["Percentage"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }


            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        public ActionResult Convert5(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtFrom"] != null && Session["dtTo"] != null)
            {
                String sql = "SELECT * FROM dbo.fn_view5_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Project_Type = rdr["Project Type"].ToString();
                            obj.Huntington = (int)rdr["Huntington"];
                            obj.Charleston = (int)rdr["Charleston"];
                            obj.Lewisburg = (int)rdr["Lewisburg"];
                            obj.Total = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view5_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model1 = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Delivery_Type = rdr["Delivery Type"].ToString();
                            obj.Huntington2 = (int)rdr["Huntington"];
                            obj.Charleston2 = (int)rdr["Charleston"];
                            obj.Lewisburg2 = (int)rdr["Lewisburg"];
                            obj.Total2 = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View5Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }
            else
            {
                String sql = "SELECT * FROM dbo.fn_view5_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Project_Type = rdr["Project Type"].ToString();
                            obj.Huntington = (int)rdr["Huntington"];
                            obj.Charleston = (int)rdr["Charleston"];
                            obj.Lewisburg = (int)rdr["Lewisburg"];
                            obj.Total = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view5_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model1 = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Delivery_Type = rdr["Delivery Type"].ToString();
                            obj.Huntington2 = (int)rdr["Huntington"];
                            obj.Charleston2 = (int)rdr["Charleston"];
                            obj.Lewisburg2 = (int)rdr["Lewisburg"];
                            obj.Total2 = (int)rdr["Total"];

                            model.Add(obj);
                        }

                    }
                    ViewBag.Print = model;
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View5Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }



        }
        public ActionResult Convert5Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtFrom"] != null && Session["dtTo"] != null)
            {
                String sql = "SELECT * FROM dbo.fn_view5_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Project_Type = rdr["Project Type"].ToString();
                            obj.Huntington = (int)rdr["Huntington"];
                            obj.Charleston = (int)rdr["Charleston"];
                            obj.Lewisburg = (int)rdr["Lewisburg"];
                            obj.Total = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view5_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model1 = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Delivery_Type = rdr["Delivery Type"].ToString();
                            obj.Huntington2 = (int)rdr["Huntington"];
                            obj.Charleston2 = (int)rdr["Charleston"];
                            obj.Lewisburg2 = (int)rdr["Lewisburg"];
                            obj.Total2 = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                }
            }
            else
            {
                String sql = "SELECT * FROM dbo.fn_view5_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Project_Type = rdr["Project Type"].ToString();
                            obj.Huntington = (int)rdr["Huntington"];
                            obj.Charleston = (int)rdr["Charleston"];
                            obj.Lewisburg = (int)rdr["Lewisburg"];
                            obj.Total = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                    gv.DataSource = model;
                    gv.DataBind();
                    Response.ClearContent();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                    Response.ContentType = "application/ms-excel";
                    Response.Charset = "";
                    StringWriter objStringWriter = new StringWriter();
                    HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                    gv.RenderControl(objHtmlTextWriter);
                    Response.Output.Write(objStringWriter.ToString());
                    Response.Flush();
                    Response.End();

                }
                sql = "SELECT * FROM dbo.fn_view5_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model1 = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (rdr["Huntington"] != System.DBNull.Value)
                        {
                            var obj = new CLSA5_1();
                            obj.Delivery_Type = rdr["Delivery Type"].ToString();
                            obj.Huntington2 = (int)rdr["Huntington"];
                            obj.Charleston2 = (int)rdr["Charleston"];
                            obj.Lewisburg2 = (int)rdr["Lewisburg"];
                            obj.Total2 = (int)rdr["Total"];

                            model.Add(obj);
                        }
                    }
                    conn.Close();
                    gv.DataSource = model;
                    gv.DataBind();
                    Response.ClearContent();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                    Response.ContentType = "application/ms-excel";
                    Response.Charset = "";
                    StringWriter objStringWriter = new StringWriter();
                    HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                    gv.RenderControl(objHtmlTextWriter);
                    Response.Output.Write(objStringWriter.ToString());
                    Response.Flush();
                    Response.End();

                }

            }
            return View();
        }
        public List<CLSA5_1> GetCLSA5_1()
        {
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "select * from dbo.fn_view5_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA5_1>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Huntington"] != System.DBNull.Value)
                    {
                        var obj = new CLSA5_1();
                        obj.Project_Type = rdr["Project Type"].ToString();
                        obj.Huntington = (int)rdr["Huntington"];
                        obj.Charleston = (int)rdr["Charleston"];
                        obj.Lewisburg = (int)rdr["Lewisburg"];
                        obj.Total = (int)rdr["Total"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }
            return model;
        }
        public ActionResult Convert6(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view6('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA6>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Designer"] != System.DBNull.Value && rdr["Designer"].ToString() != "")
                    {
                        var obj = new CLSA6();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Total_Open = (double)rdr["Total Open"];
                        obj.Total_Sold = (double)rdr["Total Sold"];
                        obj.Total_Deferred = (double)rdr["Total Deferred"];
                        obj.Total_Lost_Price = (double)rdr["Total Lost Price"];
                        obj.Total_Lost_Comp = (double)rdr["Total Lost Comp"];
                        obj.Total_Lost_Other = (double)rdr["Total Lost Other"];
                        obj.Total = (double)rdr["Total"];
                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View6Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert6Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view6('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA6>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Designer"] != System.DBNull.Value && rdr["Designer"].ToString() != "")
                    {
                        var obj = new CLSA6();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Total_Open = (double)rdr["Total Open"];
                        obj.Total_Sold = (double)rdr["Total Sold"];
                        obj.Total_Deferred = (double)rdr["Total Deferred"];
                        obj.Total_Lost_Price = (double)rdr["Total Lost Price"];
                        obj.Total_Lost_Comp = (double)rdr["Total Lost Comp"];
                        obj.Total_Lost_Other = (double)rdr["Total Lost Other"];
                        obj.Total = (double)rdr["Total"];
                        model.Add(obj);
                    }


                }
                conn.Close();
            }


            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert7(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view7('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA7>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Total Open"] != System.DBNull.Value)
                    {
                        var obj = new CLSA7();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Total_Open = (int)rdr["Total Open"];
                        obj.Total_Sold = (int)rdr["Total Sold"];
                        obj.Total_Deferred = (int)rdr["Total Deferred"];
                        obj.Total_Lost_Price = (int)rdr["Total Lost Price"];
                        obj.Total_Lost_Comp = (int)rdr["Total Lost Comp"];
                        obj.Total_Lost_Other = (int)rdr["Total Lost Other"];
                        obj.Total_Closed = (int)rdr["Total Closed"];
                        obj.Total = (int)rdr["Total"];
                        obj.Closed_Percentage = (decimal)rdr["Closed Percentage"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View7Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert7Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view7('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA7>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Total Open"] != System.DBNull.Value)
                    {
                        var obj = new CLSA7();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Total_Open = (int)rdr["Total Open"];
                        obj.Total_Sold = (int)rdr["Total Sold"];
                        obj.Total_Deferred = (int)rdr["Total Deferred"];
                        obj.Total_Lost_Price = (int)rdr["Total Lost Price"];
                        obj.Total_Lost_Comp = (int)rdr["Total Lost Comp"];
                        obj.Total_Lost_Other = (int)rdr["Total Lost Other"];
                        obj.Total_Closed = (int)rdr["Total Closed"];
                        obj.Total = (int)rdr["Total"];
                        obj.Closed_Percentage = (decimal)rdr["Closed Percentage"];

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert8(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view8('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA8>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Designer"] != System.DBNull.Value)
                    {
                        var obj = new CLSA8();
                        obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Mailing_Address = rdr["Mailing Address"].ToString();
                        obj.City = rdr["City"].ToString();
                        obj.State = rdr["State"].ToString();
                        obj.ZipCode = rdr["ZipCode"].ToString();
                        obj.Primary_Phone = rdr["Primary Phone"].ToString();
                        obj.Primary_Email = rdr["Primary Email"].ToString();


                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View8Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert8Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view8('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA8>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr["Designer"] != System.DBNull.Value)
                    {
                        var obj = new CLSA8();
                        obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                        obj.Designer = rdr["Designer"].ToString();
                        obj.Mailing_Address = rdr["Mailing Address"].ToString();
                        obj.City = rdr["City"].ToString();
                        obj.State = rdr["State"].ToString();
                        obj.ZipCode = rdr["ZipCode"].ToString();
                        obj.Primary_Phone = rdr["Primary Phone"].ToString();
                        obj.Primary_Email = rdr["Primary Email"].ToString();

                        model.Add(obj);
                    }
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert9(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view9('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA9>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA9();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Mailing_Address = rdr["Mailing Address"].ToString();
                    obj.City = rdr["City"].ToString();
                    obj.State = rdr["State"].ToString();
                    obj.ZipCode = rdr["ZipCode"].ToString();
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();

                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View9Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert9Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view9('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA9>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA9();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Mailing_Address = rdr["Mailing Address"].ToString();
                    obj.City = rdr["City"].ToString();
                    obj.State = rdr["State"].ToString();
                    obj.ZipCode = rdr["ZipCode"].ToString();
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert10(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view10('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA10>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA10();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Project_Type = rdr["Project Type"].ToString();
                    obj.Project_Status = rdr["Project Status"].ToString();
                    obj.Project_Class = rdr["Project Class"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Project_Total = rdr["Project Total"].ToString();
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View10Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert10Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view10('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA10>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA10();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Project_Type = rdr["Project Type"].ToString();
                    obj.Project_Status = rdr["Project Status"].ToString();
                    obj.Project_Class = rdr["Project Class"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Project_Total = rdr["Project Total"].ToString();
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }
        public ActionResult Convert11(string documentName, string str)
        {
            // get the data
            DateTime dtFrom;
            DateTime dtTo;

            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = System.Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = new DateTime(dtTo.Year, 1, 1);
                Session["dtFrom"] = dtFrom.ToString("yyyy-MM-dd");
                String sql = "SELECT * FROM dbo.fn_view11_1('" + dtFrom + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA11>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type = rdr["Type"].ToString();
                        obj.QTY = (int)rdr["QTY"];
                        obj.Total_Amount = (double)rdr["Total Amount"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view11_2('" + dtFrom + "','" + Session["dtTo"] + "')";

                var model1 = new List<CLSA5_1>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type2 = rdr["Type"].ToString();
                        obj.QTY2 = (int)rdr["QTY"];
                        obj.Total_Amount2 = (double)rdr["Total Amount"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view11_3('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Status = rdr["Status"].ToString();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_4('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.YTD_Statistics = rdr["YTD Statistics"].ToString();
                        obj.Numerics = (double)rdr["Numerics"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View11Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }
            else
            {

                String sql = "SELECT * FROM dbo.fn_view11_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA11>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type = rdr["Type"].ToString();
                        obj.QTY = (int)rdr["QTY"];
                        obj.Total_Amount = (double)rdr["Total Amount"];
                        obj.Total1 = 1;

                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view11_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model1 = new List<CLSA11>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type2 = rdr["Type"].ToString();
                        obj.QTY2 = (int)rdr["QTY"];
                        obj.Total_Amount2 = (double)rdr["Total Amount"];
                        obj.Total2 = 2;

                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view11_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Status = rdr["Status"].ToString();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.Total3 = 3;

                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.YTD_Statistics = rdr["YTD Statistics"].ToString();
                        obj.Numerics = (double)rdr["Numerics"];
                        obj.Total4 = 4;

                        model.Add(obj);
                    }
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View11Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }

        }
        public ActionResult Convert11Excel()
        {
            var gv = new GridView();
            DateTime dtFrom;
            DateTime dtTo;

            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = System.Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = new DateTime(dtTo.Year, 1, 1);

                String sql = "SELECT * FROM dbo.fn_view11_1('" + dtFrom + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA11>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type = rdr["Type"].ToString();
                        obj.QTY = (int)rdr["QTY"];
                        obj.Total_Amount = (double)rdr["Total Amount"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_2('" + dtFrom + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type2 = rdr["Type"].ToString();
                        obj.QTY2 = (int)rdr["QTY"];
                        obj.Total_Amount2 = (double)rdr["Total Amount"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_3('" + dtFrom + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Status = rdr["Status"].ToString();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_4('" + dtFrom + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.YTD_Statistics = rdr["YTD Statistics"].ToString();
                        obj.Numerics = (double)rdr["Numerics"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            else
            {
                String sql = "SELECT * FROM dbo.fn_view11_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA11>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type = rdr["Type"].ToString();
                        obj.QTY = (int)rdr["QTY"];
                        obj.Total_Amount = (double)rdr["Total Amount"];
                        obj.Total1 = 1;

                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view11_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Type2 = rdr["Type"].ToString();
                        obj.QTY2 = (int)rdr["QTY"];
                        obj.Total_Amount2 = (double)rdr["Total Amount"];
                        obj.Total2 = 2;

                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.Status = rdr["Status"].ToString();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.Total3 = 3;


                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view11_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA11();
                        obj.YTD_Statistics = rdr["YTD Statistics"].ToString();
                        obj.Numerics = (double)rdr["Numerics"];
                        obj.Total4 = 4;


                        model.Add(obj);
                    }
                    conn.Close();
                }
                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            return View();
        }

        public ActionResult Convert12(string documentName, string str)
        {
            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtFrom"] != null && Session["dtTo"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"].ToString() != "")
            {
                String sql = "SELECT * FROM dbo.fn_view12_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State = rdr["State"].ToString();
                        obj.No_of_Leads = (int)rdr["No. of Leads"];
                        obj.No_of_Leads_Sold = (int)rdr["No. of Leads Sold"];
                        obj.Total_amount_of_Sold_Jobs = (double)rdr["Total Amount of Sold Jobs"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view12_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model1 = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY = (int)rdr["QTY"];
                        obj.City = rdr["City"].ToString();
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view12_3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.City3 = rdr["City"].ToString();
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY4 = (int)rdr["QTY"];
                        obj.City4 = rdr["City"].ToString();
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_5('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State5 = rdr["State"].ToString();
                        obj.Installed5 = (double)rdr["Installed"];
                        obj.Pickup5 = (double)rdr["Pickup"];
                        obj.Delivered5 = (double)rdr["Delivered"];
                        obj.In_City_Installed5 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed5 = (double)rdr["Out-City Installed"];
                        obj.Remodel5 = (double)rdr["Remodel"];
                        obj.New_Construction5 = (double)rdr["New Construction"];
                        obj.Builder5 = (double)rdr["Builder"];
                        obj.Total5 = 5;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_6('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.Total = rdr["Total"].ToString();
                        obj.Installed6 = (double)rdr["Installed"];
                        obj.Pickup6 = (double)rdr["Pickup"];
                        obj.Delivered6 = (double)rdr["Delivered"];
                        obj.In_City_Installed6 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed6 = (double)rdr["Out-City Installed"];
                        obj.Remodel6 = (double)rdr["Remodel"];
                        obj.New_Construction6 = (double)rdr["New Construction"];
                        obj.Builder6 = (double)rdr["Builder"];
                        obj.Total6 = 6;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View12Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }
            else
            {

                String sql = "SELECT * FROM dbo.fn_view12_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State = rdr["State"].ToString();
                        obj.No_of_Leads = (int)rdr["No. of Leads"];
                        obj.No_of_Leads_Sold = (int)rdr["No. of Leads Sold"];
                        obj.Total_amount_of_Sold_Jobs = (double)rdr["Total Amount of Sold Jobs"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view12_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model1 = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY = (int)rdr["QTY"];
                        obj.City = rdr["City"].ToString();
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view12_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.City3 = rdr["City"].ToString();
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY4 = (int)rdr["QTY"];
                        obj.City4 = rdr["City"].ToString();
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_5('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State5 = rdr["State"].ToString();
                        obj.Installed5 = (double)rdr["Installed"];
                        obj.Pickup5 = (double)rdr["Pickup"];
                        obj.Delivered5 = (double)rdr["Delivered"];
                        obj.In_City_Installed5 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed5 = (double)rdr["Out-City Installed"];
                        obj.Remodel5 = (double)rdr["Remodel"];
                        obj.New_Construction5 = (double)rdr["New Construction"];
                        obj.Builder5 = (double)rdr["Builder"];
                        obj.Total5 = 5;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_6('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.Total = rdr["Total"].ToString();
                        obj.Installed6 = (double)rdr["Installed"];
                        obj.Pickup6 = (double)rdr["Pickup"];
                        obj.Delivered6 = (double)rdr["Delivered"];
                        obj.In_City_Installed6 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed6 = (double)rdr["Out-City Installed"];
                        obj.Remodel6 = (double)rdr["Remodel"];
                        obj.New_Construction6 = (double)rdr["New Construction"];
                        obj.Builder6 = (double)rdr["Builder"];
                        obj.Total6 = 6;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View12Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }

        }
        public ActionResult Convert12Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtFrom"] != null && Session["dtTo"] != null)
            {
                String sql = "SELECT * FROM dbo.fn_view12_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                var model = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State = rdr["State"].ToString();
                        obj.No_of_Leads = (int)rdr["No. of Leads"];
                        obj.No_of_Leads_Sold = (int)rdr["No. of Leads Sold"];
                        obj.Total_amount_of_Sold_Jobs = (double)rdr["Total Amount of Sold Jobs"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY = (int)rdr["QTY"];
                        obj.City = rdr["City"].ToString();
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.City3 = rdr["City"].ToString();
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY4 = (int)rdr["QTY"];
                        obj.City4 = rdr["City"].ToString();
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_5('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State5 = rdr["State"].ToString();
                        obj.Installed5 = (double)rdr["Installed"];
                        obj.Pickup5 = (double)rdr["Pickup"];
                        obj.Delivered5 = (double)rdr["Delivered"];
                        obj.In_City_Installed5 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed5 = (double)rdr["Out-City Installed"];
                        obj.Remodel5 = (double)rdr["Remodel"];
                        obj.New_Construction5 = (double)rdr["New Construction"];
                        obj.Builder5 = (double)rdr["Builder"];
                        obj.Total5 = 5;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_6('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.Total = rdr["Total"].ToString();
                        obj.Installed6 = (double)rdr["Installed"];
                        obj.Pickup6 = (double)rdr["Pickup"];
                        obj.Delivered6 = (double)rdr["Delivered"];
                        obj.In_City_Installed6 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed6 = (double)rdr["Out-City Installed"];
                        obj.Remodel6 = (double)rdr["Remodel"];
                        obj.New_Construction6 = (double)rdr["New Construction"];
                        obj.Builder6 = (double)rdr["Builder"];
                        obj.Total6 = 6;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            else
            {
                String sql = "SELECT * FROM dbo.fn_view12_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA12>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State = rdr["State"].ToString();
                        obj.No_of_Leads = (int)rdr["No. of Leads"];
                        obj.No_of_Leads_Sold = (int)rdr["No. of Leads Sold"];
                        obj.Total_amount_of_Sold_Jobs = (double)rdr["Total Amount of Sold Jobs"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY = (int)rdr["QTY"];
                        obj.City = rdr["City"].ToString();
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY3 = (int)rdr["QTY"];
                        obj.City3 = rdr["City"].ToString();
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.QTY4 = (int)rdr["QTY"];
                        obj.City4 = rdr["City"].ToString();
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_5('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.State5 = rdr["State"].ToString();
                        obj.Installed5 = (double)rdr["Installed"];
                        obj.Pickup5 = (double)rdr["Pickup"];
                        obj.Delivered5 = (double)rdr["Delivered"];
                        obj.In_City_Installed5 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed5 = (double)rdr["Out-City Installed"];
                        obj.Remodel5 = (double)rdr["Remodel"];
                        obj.New_Construction5 = (double)rdr["New Construction"];
                        obj.Builder5 = (double)rdr["Builder"];
                        obj.Total5 = 5;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view12_6('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA12();
                        obj.Total = rdr["Total"].ToString();
                        obj.Installed6 = (double)rdr["Installed"];
                        obj.Pickup6 = (double)rdr["Pickup"];
                        obj.Delivered6 = (double)rdr["Delivered"];
                        obj.In_City_Installed6 = (double)rdr["In-City Installed"];
                        obj.Out_City_Installed6 = (double)rdr["Out-City Installed"];
                        obj.Remodel6 = (double)rdr["Remodel"];
                        obj.New_Construction6 = (double)rdr["New Construction"];
                        obj.Builder6 = (double)rdr["Builder"];
                        obj.Total6 = 6;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            return View();
        }

        public ActionResult Convert13(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view13('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA13>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA13();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Sold_Date = (DateTime)rdr["Sold Date"];
                    obj.Responsible_Party = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Project_Name = rdr["Project Name"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Installed_Total = (double)rdr["Installed Total"];
                    obj.Delivered_Total_Before_Taxes = (double)rdr["Delivered Total Before Taxes"];


                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View13Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert13Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view13('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA13>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA13();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Sold_Date = (DateTime)rdr["Sold Date"];
                    obj.Responsible_Party = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Project_Name = rdr["Project Name"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Installed_Total = (double)rdr["Installed Total"];
                    obj.Delivered_Total_Before_Taxes = (double)rdr["Delivered Total Before Taxes"];


                    model.Add(obj);
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert14(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view14('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA14>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA14();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Assigned_Leads = (int)rdr["No. of all assigned Leads"];
                    obj.Total_amount_all_leads = (double)rdr["Total amount of all leads assigned"];
                    obj.Sold_Jobs_Only = (int)rdr["No. of Sold Jobs Only"];
                    obj.Total_amount_Sold_Jobs = (double)rdr["Total amount of Sold Jobs Only"];
                    obj.Lost_Jobs_Only = (int)rdr["No. of Lost Jobs Only"];
                    obj.Total_Amount_Lost_Jobs = (double)rdr["Total amount of Lost Jobs"];
                    obj.Closed_Percentage = (decimal)rdr["Closed Percentage"];
                    obj.Avg_Days_Sell = rdr["Avg Days to Sell a Lead"].ToString();
                    obj.Avg_Amount_Sold_Jobs = rdr["Avg Amount of Sold Jobs"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View14Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert14Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view14('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA14>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA14();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Assigned_Leads = (int)rdr["No. of all assigned Leads"];
                    obj.Total_amount_all_leads = (double)rdr["Total amount of all leads assigned"];
                    obj.Sold_Jobs_Only = (int)rdr["No. of Sold Jobs Only"];
                    obj.Total_amount_Sold_Jobs = (double)rdr["Total amount of Sold Jobs Only"];
                    obj.Lost_Jobs_Only = (int)rdr["No. of Lost Jobs Only"];
                    obj.Total_Amount_Lost_Jobs = (double)rdr["Total amount of Lost Jobs"];
                    obj.Closed_Percentage = (decimal)rdr["Closed Percentage"];
                    obj.Avg_Days_Sell = rdr["Avg Days to Sell a Lead"].ToString();
                    obj.Avg_Amount_Sold_Jobs = rdr["Avg Amount of Sold Jobs"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert15(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view15('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA15>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA15();
                    obj.Month = rdr["Month"].ToString();
                    obj.Huntington = (double)rdr["Huntington"];
                    obj.Charleston = (double)rdr["Charleston"];
                    obj.Lewisburg = (double)rdr["Lewisburg"];
                    obj.Total = (double)rdr["Total"];
                    obj.Percentage = (decimal)rdr["Percentage"];

                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View15Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert15Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view15('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA15>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA15();
                    obj.Month = rdr["Month"].ToString();
                    obj.Huntington = (double)rdr["Huntington"];
                    obj.Charleston = (double)rdr["Charleston"];
                    obj.Lewisburg = (double)rdr["Lewisburg"];
                    obj.Total = (double)rdr["Total"];
                    obj.Percentage = (decimal)rdr["Percentage"];

                    model.Add(obj);
                }
                conn.Close();
            }


            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert16(string documentName, string str)
        {

            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view16('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA16>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA16();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Project_Type = rdr["Project Type"].ToString();
                    obj.Project_Class = rdr["Project Class"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Project_Total = (double)rdr["Project Total"];
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            //var lsr = new lead_source_report();
            //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // get html code from url
            string htmlString = this.RenderView("View16Print", model);


            // get base url (to resolve relative links to external resources)
            //this doesn't work with 'http' unless using the paid version
            //var uri = new Uri(url);
            //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

            // set converter options
            converter.Options.PdfPageSize = PdfPageSize.Letter;
            converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
            converter.Options.MarginLeft = 5;
            converter.Options.MarginRight = 5;
            converter.Options.MarginTop = 5;
            converter.Options.MarginBottom = 5;

            // create a new pdf document converting the html string
            PdfDocument doc = converter.ConvertHtmlString(htmlString);

            // get conversion result (contains document info from the web page)
            HtmlToPdfResult result = converter.ConversionResult;

            // set the document properties
            doc.DocumentInformation.Title = result.WebPageInformation.Title;
            doc.DocumentInformation.Subject = result.WebPageInformation.Description;
            doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

            doc.DocumentInformation.Author = "CreativeKitchens";
            doc.DocumentInformation.CreationDate = DateTime.Now;

            // save pdf document
            byte[] pdf = doc.Save();

            // close pdf document
            doc.Close();

            //return resulted pdf document
            FileResult fileResult = new FileContentResult(pdf, "application/pdf")
            {
                FileDownloadName = documentName + ".pdf"
            };
            return fileResult;
        }
        public ActionResult Convert16Excel()
        {
            var gv = new GridView();


            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            String sql = "SELECT * FROM dbo.fn_view16('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";

            var model = new List<CLSA16>();
            using (SqlConnection conn = new SqlConnection(constr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var obj = new CLSA16();
                    obj.Responsible_Party_Or_Lead_Creator = rdr["Responsible Party/Lead Creator"].ToString();
                    obj.Designer = rdr["Designer"].ToString();
                    obj.Project_Type = rdr["Project Type"].ToString();
                    obj.Project_Class = rdr["Project Class"].ToString();
                    obj.Delivery_Status = rdr["Delivery Status"].ToString();
                    obj.Project_Total = (double)rdr["Project Total"];
                    obj.Primary_Phone = rdr["Primary Phone"].ToString();
                    obj.Primary_Email = rdr["Primary Email"].ToString();


                    model.Add(obj);
                }
                conn.Close();
            }

            gv.DataSource = model;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View();
        }

        public ActionResult Convert17(string documentName, string str)
        {
            // get the data
            DateTime dtFrom;
            DateTime dtTo;

            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = System.Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = new DateTime(dtTo.Year, 1, 1);

                String sql = "SELECT * FROM dbo.fn_view17_1('" + dtFrom + "','" + Session["dtTo"] + "')";
                var model = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Branch_Name = rdr["Branch Name"].ToString();
                        obj.YTD_Total_Sales = (double)rdr["YTD Total Sales"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view17_2('" + dtFrom + "','" + Session["dtTo"] + "')";

                var model1 = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month = rdr["Month"].ToString();
                        obj.Price = (double)rdr["Price"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view17_3('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month3 = rdr["Month"].ToString();
                        obj.Price3 = (double)rdr["Price"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_4('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month4 = rdr["Month"].ToString();
                        obj.Price4 = (double)rdr["Price"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View17Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }
            else
            {

                String sql = "SELECT * FROM dbo.fn_view17_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Branch_Name = rdr["Branch Name"].ToString();
                        obj.YTD_Total_Sales = (double)rdr["YTD Total Sales"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view17_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";

                var model1 = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month = rdr["Month"].ToString();
                        obj.Price = (double)rdr["Price"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                sql = "SELECT * FROM dbo.fn_view17_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month3 = rdr["Month"].ToString();
                        obj.Price3 = (double)rdr["Price"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month4 = rdr["Month"].ToString();
                        obj.Price4 = (double)rdr["Price"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                //var lsr = new lead_source_report();
                //PagedList<CLSA> model1 = new PagedList<CLSA>(null, 1, 3);

                // instantiate a html to pdf converter object
                HtmlToPdf converter = new HtmlToPdf();

                // get html code from url
                string htmlString = this.RenderView("View17Print", model);


                // get base url (to resolve relative links to external resources)
                //this doesn't work with 'http' unless using the paid version
                //var uri = new Uri(url);
                //string baseUrl = uri.GetLeftPart(System.UriPartial.Authority);

                // set converter options
                converter.Options.PdfPageSize = PdfPageSize.Letter;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                converter.Options.MarginLeft = 5;
                converter.Options.MarginRight = 5;
                converter.Options.MarginTop = 5;
                converter.Options.MarginBottom = 5;

                // create a new pdf document converting the html string
                PdfDocument doc = converter.ConvertHtmlString(htmlString);

                // get conversion result (contains document info from the web page)
                HtmlToPdfResult result = converter.ConversionResult;

                // set the document properties
                doc.DocumentInformation.Title = result.WebPageInformation.Title;
                doc.DocumentInformation.Subject = result.WebPageInformation.Description;
                doc.DocumentInformation.Keywords = result.WebPageInformation.Keywords;

                doc.DocumentInformation.Author = "CreativeKitchens";
                doc.DocumentInformation.CreationDate = DateTime.Now;

                // save pdf document
                byte[] pdf = doc.Save();

                // close pdf document
                doc.Close();

                //return resulted pdf document
                FileResult fileResult = new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = documentName + ".pdf"
                };
                return fileResult;
            }

        }
        public ActionResult Convert17Excel()
        {
            var gv = new GridView();
            // get the data
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            if (Session["dtFrom"] != null && Session["dtTo"] != null)
            {
                String sql = "SELECT * FROM dbo.fn_view17_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                var model = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Branch_Name = rdr["Branch Name"].ToString();
                        obj.YTD_Total_Sales = (double)rdr["YTD Total Sales"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month = rdr["Month"].ToString();
                        obj.Price = (double)rdr["Price"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month3 = rdr["Month"].ToString();
                        obj.Price3 = (double)rdr["Price"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month4 = rdr["Month"].ToString();
                        obj.Price4 = (double)rdr["Price"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            else
            {
                String sql = "SELECT * FROM dbo.fn_view17_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                var model = new List<CLSA17>();
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Branch_Name = rdr["Branch Name"].ToString();
                        obj.YTD_Total_Sales = (double)rdr["YTD Total Sales"];
                        obj.Total1 = 1;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month = rdr["Month"].ToString();
                        obj.Price = (double)rdr["Price"];
                        obj.Total2 = 2;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month3 = rdr["Month"].ToString();
                        obj.Price3 = (double)rdr["Price"];
                        obj.Total3 = 3;
                        model.Add(obj);
                    }
                    conn.Close();
                }
                sql = "SELECT * FROM dbo.fn_view17_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection conn = new SqlConnection(constr))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var obj = new CLSA17();
                        obj.Month4 = rdr["Month"].ToString();
                        obj.Price4 = (double)rdr["Price"];
                        obj.Total4 = 4;
                        model.Add(obj);
                    }
                    conn.Close();
                }

                gv.DataSource = model;
                gv.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=CLSA.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
            }
            return View();
        }
    }
}