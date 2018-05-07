using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using PagedList.Mvc;
using System.Configuration;
using System.Data.SqlClient;
using ck_project.Models;
using System.Data;
using System.Dynamic;

namespace ck_project.Controllers
{
    public class ReportsController : Controller
    {
        // GET: Reports
        ckdatabase db = new ckdatabase();

        public ActionResult CompanyLeadSourceAnalysis(int? page, string search, string search1)
        {
            string query;
            string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            SqlConnection con = new SqlConnection(constr);
            //DataTable dt = new DataTable();

            query = "ALTER view [dbo].[reported] as select lead_source.source_name,branch_name,count(leads.branch_number) Branch_Count from leads,branches,lead_source where leads.branch_number=branches.branch_number and leads.source_number=lead_source.source_number";
            if (search != null && search1 != null)
            {
                query = query + " and lead_date>='" + search + "' and lead_date<='" + search1 + "'";
            }
            else
            {
                query = query + " and lead_date>='" + "01-01-1900" + "' and lead_date<='" + "01-01-1900" + "'";
            }

            query = query + " group by lead_source.source_name,branch_name";

            Session["dtFrom"] = search;
            Session["dtTo"] = search1;

            var ctx1 = new ckdatabase();
            int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);


            return View(db.lead_source_report.Where(x => x.total > 0).ToList().ToPagedList(page ?? 1, 100));
        }
        public ActionResult CompanyProjectTypeAnalysis(int? page, string search, string search1)
        {
            //string query="";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);
            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view2('" + search + "','" + search1 + "')";

            //    var ctx = new ckdatabase();
            //    int noOfRowUpdated = ctx.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx = new ckdatabase();
            //    int noOfRowUpdated = ctx.Database.ExecuteSqlCommand(query);
            //}
            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view2(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view2(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}

            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA2 = GetCLSA2();
            return View(model);

        }
        public List<CLSA2> GetCLSA2()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA2> Data = new List<CLSA2>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view2('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Huntington"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA2
                                        {
                                            ProjectName = sdr["Project Type Name"].ToString(),
                                            Huntington = (Int32)sdr["Huntington"],
                                            Charleston = (Int32)sdr["Charleston"],
                                            Lewisburg = (Int32)sdr["Lewisburg"],
                                            total = (Int32)sdr["Total"]

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult CompanyLeadStatusAnalysis(int? page, string search, string search1)
        {
            //string query="";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            // if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view3('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view3(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view3(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA3 = GetCLSA3();
            return View(model);
        }

        public List<CLSA3> GetCLSA3()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA3> Data = new List<CLSA3>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view3('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Huntington"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA3
                                        {
                                            Project_Status_Name = sdr["Project Status Name"].ToString(),
                                            Huntington = (int)sdr["Huntington"],
                                            Charleston = (int)sdr["Charleston"],
                                            Lewisburg = (int)sdr["Lewisburg"],
                                            total = (int)sdr["Total"]

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;


        }
        public ActionResult CompanyProductTypeAnalysis(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view4('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view4(From, To).Where(x => x.Company_Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view4(From, To).Where(x => x.Company_Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA4 = GetCLSA4();
            return View(model);
        }
        public List<CLSA4> GetCLSA4()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA4> Data = new List<CLSA4>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view4('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Huntington"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA4
                                        {
                                            Type = sdr["Types"].ToString(),
                                            Huntington = (double)sdr["Huntington"],
                                            Charleston = (double)sdr["Charleston"],
                                            Lewisburg = (double)sdr["Lewisburg"],
                                            Companytotal = (double)sdr["Company Total"],
                                            Percentage = (double)sdr["Percentage"]

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult CompanyLeadTypeCategoryAnalysis(int? page, string search, string search1)
        {
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA5_1S = GetCLSA5_1();

            return View(model);
        }

        public List<CLSA5_1> GetCLSA5_1()
        {
            string query = "";
            List<CLSA5_1> ProjectType = new List<CLSA5_1>();
            if (Session["dtFrom"] != null && Session["dtTo"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"].ToString() != "")
            {

                query = "select * from dbo.fn_view5_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    ProjectType.Add(new CLSA5_1
                                    {
                                        Project_Type = sdr["Project Type"].ToString(),
                                        Huntington = (int)sdr["Huntington"],
                                        Charleston = (int)sdr["Charleston"],
                                        Lewisburg = (int)sdr["Lewisburg"],
                                        Total = (int)sdr["Total"]
                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }
            }
            if (Session["dtFrom"] != null && Session["dtTo"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"].ToString() != "")
            {
                query = "select * from dbo.fn_view5_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    ProjectType.Add(new CLSA5_1
                                    {
                                        Delivery_Type = sdr["Delivery Type"].ToString(),
                                        Huntington2 = (int)sdr["Huntington"],
                                        Charleston2 = (int)sdr["Charleston"],
                                        Lewisburg2 = (int)sdr["Lewisburg"],
                                        Total2 = (int)sdr["Total"]
                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }

            else
            {
                query = "select * from dbo.fn_view5_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    ProjectType.Add(new CLSA5_1
                                    {
                                        Project_Type = sdr["Project Type"].ToString(),
                                        Huntington = (int)sdr["Huntington"],
                                        Charleston = (int)sdr["Charleston"],
                                        Lewisburg = (int)sdr["Lewisburg"],
                                        Total = (int)sdr["Total"]
                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }
                query = "select * from dbo.fn_view5_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    ProjectType.Add(new CLSA5_1
                                    {
                                        Delivery_Type = sdr["Delivery Type"].ToString(),
                                        Huntington2 = (int)sdr["Huntington"],
                                        Charleston2 = (int)sdr["Charleston"],
                                        Lewisburg2 = (int)sdr["Lewisburg"],
                                        Total2 = (int)sdr["Total"]
                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }
            return ProjectType;
        }

        public List<CLSA5_2> GetCLSA5_2()
        {
            string query = "";
            List<CLSA5_2> DeliveyType = new List<CLSA5_2>();
            if (Session["dtFrom"] != null && Session["dtTo"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"].ToString() != "")
            {
                query = "select * from dbo.fn_view5_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    DeliveyType.Add(new CLSA5_2
                                    {
                                        Delivery_Type = sdr["Delivery Type"].ToString(),
                                        Huntington = (int)sdr["Huntington"],
                                        Charleston = (int)sdr["Charleston"],
                                        Lewisburg = (int)sdr["Lewisburg"],
                                        Total = (int)sdr["Total"]
                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }
            else
            {
                query = "select * from dbo.fn_view5_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Huntington"] != System.DBNull.Value)
                                {
                                    DeliveyType.Add(new CLSA5_2
                                    {
                                        Delivery_Type = sdr["Delivery Type"].ToString(),
                                        Huntington = (int)sdr["Huntington"],
                                        Charleston = (int)sdr["Charleston"],
                                        Lewisburg = (int)sdr["Lewisburg"],
                                        Total = (int)sdr["Total"]
                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }
            return DeliveyType;
        }

        public ActionResult CompanyDesignerLeadStatusTotals(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view6('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view6('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view6(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view6(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA6 = GetCLSA6();
            return View(model);
        }
        public List<CLSA6> GetCLSA6()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA6> Data = new List<CLSA6>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view6('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Designer"] != System.DBNull.Value && sdr["Designer"].ToString() != "")
                                    {
                                        Data.Add(new CLSA6
                                        {
                                            Designer = sdr["Designer"].ToString(),
                                            Total_Open = (double)sdr["Total Open"],
                                            Total_Sold = (double)sdr["Total Sold"],
                                            Total_Deferred = (double)sdr["Total Deferred"],
                                            Total_Lost_Price = (double)sdr["Total Lost Price"],
                                            Total_Lost_Comp = (double)sdr["Total Lost Comp"],
                                            Total_Lost_Other = (double)sdr["Total Lost Other"],
                                            Total = (double)sdr["Total"]

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }
        public ActionResult CompanyDesignerLeadStatus(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view7('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view7('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view7(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view7(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA7 = GetCLSA7();
            return View(model);

        }
        public List<CLSA7> GetCLSA7()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA7> Data = new List<CLSA7>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view7('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Total Open"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA7
                                        {
                                            Designer = sdr["Designer"].ToString(),
                                            Total_Open = (Int32)sdr["Total Open"],
                                            Total_Sold = (Int32)sdr["Total Sold"],
                                            Total_Deferred = (Int32)sdr["Total Deferred"],
                                            Total_Lost_Price = (Int32)sdr["Total Lost Price"],
                                            Total_Lost_Comp = (Int32)sdr["Total Lost Comp"],
                                            Total_Lost_Other = (Int32)sdr["Total Lost Other"],
                                            Total_Closed = (Int32)sdr["Total Closed"],
                                            Total = (Int32)sdr["Total"],
                                            Closed_Percentage = (decimal)sdr["Closed Percentage"]

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult CompanySoldAndClosedJobs(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view8('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view8('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view8(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view8(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA8 = GetCLSA8();
            return View(model);
        }

        public List<CLSA8> GetCLSA8()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA8> Data = new List<CLSA8>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view8('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Designer"] != System.DBNull.Value && sdr["Designer"].ToString() != "")
                                    {
                                        Data.Add(new CLSA8
                                        {
                                            Responsible_Party_Or_Lead_Creator = sdr["Responsible Party/Lead Creator"].ToString(),
                                            Designer = sdr["Designer"].ToString(),
                                            Mailing_Address = sdr["Mailing Address"].ToString(),
                                            City = sdr["City"].ToString(),
                                            State = sdr["State"].ToString(),
                                            ZipCode = sdr["ZipCode"].ToString(),
                                            Primary_Phone = sdr["Primary Phone"].ToString(),
                                            Primary_Email = sdr["Primary Email"].ToString()


                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult CompanySoldJobsOnly(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view9('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view9('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view9(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view9(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA9 = GetCLSA9();
            return View(model);

        }
        public List<CLSA9> GetCLSA9()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA9> Data = new List<CLSA9>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view9('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Designer"] != System.DBNull.Value && sdr["Designer"].ToString() != "")
                                    {
                                        Data.Add(new CLSA9
                                        {
                                            Responsible_Party_Or_Lead_Creator = sdr["Responsible Party/Lead Creator"].ToString(),
                                            Designer = sdr["Designer"].ToString(),
                                            Mailing_Address = sdr["Mailing Address"].ToString(),
                                            City = sdr["City"].ToString(),
                                            State = sdr["State"].ToString(),
                                            ZipCode = sdr["ZipCode"].ToString(),
                                            Primary_Phone = sdr["Primary Phone"].ToString(),
                                            Primary_Email = sdr["Primary Email"].ToString()


                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult DesignerActiveLeadReport(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view10('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view10('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{

            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view10(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view10(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA10 = GetCLSA10();
            return View(model);
        }
        public List<CLSA10> GetCLSA10()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA10> Data = new List<CLSA10>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view10('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Project Type"] != System.DBNull.Value && sdr["Project Type"].ToString() != "")
                                    {
                                        Data.Add(new CLSA10
                                        {
                                            Responsible_Party_Or_Lead_Creator = sdr["Responsible Party/Lead Creator"].ToString(),
                                            Project_Type = sdr["Project Type"].ToString(),
                                            Project_Status = sdr["Project Status"].ToString(),
                                            Project_Class = sdr["Project Class"].ToString(),
                                            Delivery_Status = sdr["Delivery Status"].ToString(),
                                            Project_Total = sdr["Project Total"].ToString(),
                                            Primary_Phone = sdr["Primary Phone"].ToString(),
                                            Primary_Email = sdr["Primary Email"].ToString()


                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }
        public ActionResult DesignerYearToDateReport(int? page, string search, string search1)
        {
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA11 = GetCLSA11();
            return View(model);
        }

        public List<CLSA11> GetCLSA11()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;

            List<CLSA11> Data = new List<CLSA11>();
            if (Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = new DateTime(dtTo.Year, 1, 1);
                Session["dtFrom"] = dtFrom.ToString();

                query = "select * from dbo.fn_view11_1('" + dtFrom + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA11
                                    {
                                        Type = sdr["Type"].ToString(),
                                        QTY = (int)sdr["QTY"],
                                        Total_Amount = (double)sdr["Total Amount"],
                                        Total1 = 1,

                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }

                query = "select * from dbo.fn_view11_2('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA11
                                    {
                                        Type2 = sdr["Type"].ToString(),
                                        QTY2 = (int)sdr["QTY"],
                                        Total_Amount2 = (double)sdr["Total Amount"],
                                        Total2 = 2,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view11_3('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA11
                                    {
                                        Status = sdr["Status"].ToString(),
                                        QTY3 = (int)sdr["QTY"],
                                        Total3 = 3,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view11_4('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Numerics"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA11
                                    {
                                        YTD_Statistics = sdr["YTD Statistics"].ToString(),
                                        Numerics = (double)sdr["Numerics"],
                                        Total4 = 4,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }



            return Data;
        }

        public ActionResult LocationStatisticsOfSoldJobs(int? page, string search, string search1)
        {
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA12 = GetCLSA12();
            return View(model);
        }

        public List<CLSA12> GetCLSA12()
        {
            string query = "";
            List<CLSA12> Data = new List<CLSA12>();
            if (Session["dtFrom"] != null && Session["dtTo"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"].ToString() != "")
            {

                query = "select * from dbo.fn_view12_1('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["No. of Leads"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        State = sdr["State"].ToString(),
                                        No_of_Leads = (int)sdr["No. of Leads"],
                                        No_of_Leads_Sold = (int)sdr["No. of Leads Sold"],
                                        Total_amount_of_Sold_Jobs = (double)sdr["Total amount of Sold Jobs"],
                                        Total1 = 1,

                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }

                query = "select * from dbo.fn_view12_2('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY = (int)sdr["QTY"],
                                        City = sdr["City"].ToString(),
                                        Total2 = 2,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_3('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY3 = (int)sdr["QTY"],
                                        City3 = sdr["City"].ToString(),
                                        Total3 = 3,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_4('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY4 = (int)sdr["QTY"],
                                        City4 = sdr["City"].ToString(),
                                        Total4 = 4,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_5('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Installed"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        State5 = sdr["State"].ToString(),
                                        Installed5 = (double)sdr["Installed"],
                                        Pickup5 = (double)sdr["Pickup"],
                                        Delivered5 = (double)sdr["Delivered"],
                                        In_City_Installed5 = (double)sdr["In-City Installed"],
                                        Out_City_Installed5 = (double)sdr["Out-City Installed"],
                                        Remodel5 = (double)sdr["Remodel"],
                                        New_Construction5 = (double)sdr["New Construction"],
                                        Builder5 = (double)sdr["Builder"],
                                        Total5 = 5,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_6('" + Session["dtFrom"] + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Installed"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        Total = sdr["Total"].ToString(),
                                        Installed6 = (double)sdr["Installed"],
                                        Pickup6 = (double)sdr["Pickup"],
                                        Delivered6 = (double)sdr["Delivered"],
                                        In_City_Installed6 = (double)sdr["In-City Installed"],
                                        Out_City_Installed6 = (double)sdr["Out-City Installed"],
                                        Remodel6 = (double)sdr["Remodel"],
                                        New_Construction6 = (double)sdr["New Construction"],
                                        Builder6 = (double)sdr["Builder"],
                                        Total6 = 6,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }

            else
            {
                query = "select * from dbo.fn_view12_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["No. of Leads"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        State = sdr["State"].ToString(),
                                        No_of_Leads = (int)sdr["No. of Leads"],
                                        No_of_Leads_Sold = (int)sdr["No. of Leads Sold"],
                                        Total_amount_of_Sold_Jobs = (double)sdr["Total amount of Sold Jobs"],
                                        Total1 = 1,

                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }

                query = "select * from dbo.fn_view12_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY = (int)sdr["QTY"],
                                        City = sdr["City"].ToString(),
                                        Total2 = 2,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY3 = (int)sdr["QTY"],
                                        City3 = sdr["City"].ToString(),
                                        Total3 = 3,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["QTY"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        QTY4 = (int)sdr["QTY"],
                                        City4 = sdr["City"].ToString(),
                                        Total4 = 4,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_5('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Installed"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        State5 = sdr["State"].ToString(),
                                        Installed5 = (double)sdr["Installed"],
                                        Pickup5 = (double)sdr["Pickup"],
                                        Delivered5 = (double)sdr["Delivered"],
                                        In_City_Installed5 = (double)sdr["In-City Installed"],
                                        Out_City_Installed5 = (double)sdr["Out-City Installed"],
                                        Remodel5 = (double)sdr["Remodel"],
                                        New_Construction5 = (double)sdr["New Construction"],
                                        Builder5 = (double)sdr["Builder"],
                                        Total5 = 5,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view12_6('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Installed"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA12
                                    {
                                        Total = sdr["Total"].ToString(),
                                        Installed6 = (double)sdr["Installed"],
                                        Pickup6 = (double)sdr["Pickup"],
                                        Delivered6 = (double)sdr["Delivered"],
                                        In_City_Installed6 = (double)sdr["In-City Installed"],
                                        Out_City_Installed6 = (double)sdr["Out-City Installed"],
                                        Remodel6 = (double)sdr["Remodel"],
                                        New_Construction6 = (double)sdr["New Construction"],
                                        Builder6 = (double)sdr["Builder"],
                                        Total6 = 6,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }
            return Data;
        }

        public ActionResult AccountingSoldJobReport(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view13('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view13('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view13(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view13(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA13 = GetCLSA13();
            return View(model);
        }
        public List<CLSA13> GetCLSA13()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA13> Data = new List<CLSA13>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view13('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Installed Total"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA13
                                        {
                                            Designer = sdr["Designer"].ToString(),
                                            Sold_Date = (DateTime)sdr["Sold Date"],
                                            Responsible_Party = sdr["Responsible Party/Lead Creator"].ToString(),
                                            Project_Name = sdr["Project Name"].ToString(),
                                            Delivery_Status = sdr["Delivery Status"].ToString(),
                                            Installed_Total = (double)sdr["Installed Total"],
                                            Delivered_Total_Before_Taxes = (double)sdr["Delivered Total Before Taxes"],

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }
        public ActionResult CompanyYearToDateByDesignerReport(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view14('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view14('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view14(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view14(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA14 = GetCLSA14();
            return View(model);
        }
        public List<CLSA14> GetCLSA14()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA14> Data = new List<CLSA14>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view14('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Total amount of Lost Jobs"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA14
                                        {
                                            Designer = sdr["Designer"].ToString(),
                                            Assigned_Leads = (int)sdr["No. of all assigned Leads"],
                                            Total_amount_all_leads = (double)sdr["Total amount of all leads assigned"],
                                            Sold_Jobs_Only = (int)sdr["No. of Sold Jobs Only"],
                                            Total_amount_Sold_Jobs = (double)sdr["Total amount of Sold Jobs Only"],
                                            Lost_Jobs_Only = (int)sdr["No. of Lost Jobs Only"],
                                            Total_Amount_Lost_Jobs = (double)sdr["Total amount of Lost Jobs"],
                                            Closed_Percentage = (decimal)sdr["Closed Percentage"],
                                            Avg_Days_Sell = sdr["Avg Days to Sell a Lead"].ToString(),
                                            Avg_Amount_Sold_Jobs = sdr["Avg Amount of Sold Jobs"].ToString(),

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult MonthlySalesReport(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view15('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view15('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view15(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view15(From, To).Where(x => x.Total > 0).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA15 = GetCLSA15();
            return View(model);

        }

        public List<CLSA15> GetCLSA15()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA15> Data = new List<CLSA15>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view15('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Huntington"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA15
                                        {
                                            Month = sdr["Month"].ToString(),
                                            Huntington = (double)sdr["Huntington"],
                                            Charleston = (double)sdr["Charleston"],
                                            Lewisburg = (double)sdr["Lewisburg"],
                                            Total = (double)sdr["Total"],
                                            Percentage = (decimal)sdr["Percentage"],

                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult AllOpenLeads(int? page, string search, string search1)
        {
            //string query = "";
            //string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
            //SqlConnection con = new SqlConnection(constr);

            //if (search != null && search1 != null)
            //{

            //    query = " select * from dbo.fn_view16('" + search + "','" + search1 + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}
            //else
            //{

            //    query = " select * from dbo.fn_view16('" + "01-01-1900" + "','" + "01-01-1900" + "')";
            //    var ctx1 = new ckdatabase();
            //    int noOfRowUpdated = ctx1.Database.ExecuteSqlCommand(query);
            //}

            //Session["dtFrom"] = search;
            //Session["dtTo"] = search1;
            //if (search != "" && search != null)
            //{
            //    DateTime From = Convert.ToDateTime(search);
            //    DateTime To = Convert.ToDateTime(search1);
            //    return View(db.fn_view16(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            //else
            //{
            //    DateTime From = Convert.ToDateTime("01-01-1900");
            //    DateTime To = Convert.ToDateTime("01-01-1900");
            //    return View(db.fn_view16(From, To).ToList().ToPagedList(page ?? 1, 100));
            //}
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA16 = GetCLSA16();
            return View(model);
        }

        public List<CLSA16> GetCLSA16()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA16> Data = new List<CLSA16>();
            if (Session["dtFrom"] != null && Session["dtFrom"].ToString() != "" && Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = Convert.ToDateTime(Session["dtFrom"].ToString());
                query = "select * from dbo.fn_view16('" + dtFrom + "','" + dtTo + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            if (sdr.HasRows == true)
                            {
                                while (sdr.Read())
                                {
                                    if (sdr["Designer"] != System.DBNull.Value)
                                    {
                                        Data.Add(new CLSA16
                                        {
                                            Responsible_Party_Or_Lead_Creator = sdr["Responsible Party/Lead Creator"].ToString(),
                                            Designer = sdr["Designer"].ToString(),
                                            Project_Type = sdr["Project Type"].ToString(),
                                            Project_Class = sdr["Project Class"].ToString(),
                                            Delivery_Status = sdr["Delivery Status"].ToString(),
                                            Project_Total = (double)sdr["Project Total"],
                                            Primary_Phone = sdr["Primary Phone"].ToString(),
                                            Primary_Email = sdr["Primary Email"].ToString(),
                                        });
                                    }
                                }
                            }
                        }
                        con.Close();

                    }

                }

            }
            return Data;
        }

        public ActionResult CompanySalesDashboard(int? page, string search, string search1)
        {
            Session["dtFrom"] = search;
            Session["dtTo"] = search1;
            dynamic model = new ExpandoObject();
            model.CLSA17 = GetCLSA17();
            return View(model);
        }

        public List<CLSA17> GetCLSA17()
        {
            string query = "";
            DateTime dtFrom;
            DateTime dtTo;
            List<CLSA17> Data = new List<CLSA17>();
            if (Session["dtTo"] != null && Session["dtTo"].ToString() != "")
            {
                dtTo = Convert.ToDateTime(Session["dtTo"].ToString());
                dtFrom = new DateTime(dtTo.Year, 1, 1);
                Session["dtFrom"] = dtFrom;
                query = "select * from dbo.fn_view17_1('" + dtFrom + "','" + Session["dtTo"] + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["YTD Total Sales"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Branch_Name = sdr["Branch Name"].ToString(),
                                        YTD_Total_Sales = (double)sdr["YTD Total Sales"],
                                        Total1 = 1,

                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }

                query = "select * from dbo.fn_view17_2('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month = sdr["Month"].ToString(),
                                        Price = (double)sdr["Price"],
                                        Total2 = 2,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view17_3('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month3 = sdr["Month"].ToString(),
                                        Price3 = (double)sdr["Price"],
                                        Total3 = 3,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view17_4('" + dtFrom + "','" + Session["dtTo"] + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())

                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month4 = sdr["Month"].ToString(),
                                        Price4 = (double)sdr["Price"],
                                        Total4 = 4,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }

            else
            {
                query = "select * from dbo.fn_view17_1('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                string constr = ConfigurationManager.ConnectionStrings["CKConnection"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["YTD Total Sales"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Branch_Name = sdr["Branch Name"].ToString(),
                                        YTD_Total_Sales = (double)sdr["YTD Total Sales"],
                                        Total1 = 1,

                                    });
                                }
                            }
                        }
                        con.Close();

                    }
                }

                query = "select * from dbo.fn_view17_2('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month = sdr["Month"].ToString(),
                                        Price = (double)sdr["Price"],
                                        Total2 = 2,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view17_3('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month3 = sdr["Month"].ToString(),
                                        Price3 = (double)sdr["Price"],
                                        Total3 = 3,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
                query = "select * from dbo.fn_view17_4('" + "01-01-1900" + "','" + "01-01-1900" + "')";
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        con.Open();
                        using (SqlDataReader sdr = cmd.ExecuteReader())
                        {
                            while (sdr.Read())
                            {
                                if (sdr["Price"] != System.DBNull.Value)
                                {
                                    Data.Add(new CLSA17
                                    {
                                        Month4 = sdr["Month"].ToString(),
                                        Price4 = (double)sdr["Price"],
                                        Total4 = 4,

                                    });
                                }
                            }
                            con.Close();

                        }
                    }
                }
            }
            return Data;
        }
    }
}