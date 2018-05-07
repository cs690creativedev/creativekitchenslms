using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using PagedList.Mvc;

namespace ck_project.Controllers
{
    public class CountyTaxController : Controller
    {
        //Creating the db connecton
        ckdatabase db = new ckdatabase();
        static List<tax> lst = new List<tax>();
        // GET: Customers


        public ActionResult CountyTaxNew(int? page, string search = null, string Type = null, string Type1 = null, string Start = null, string end = null, string msg = null)
        {
            //DateTime start = string.IsNullOrEmpty(Start) ? DateTime.MinValue : DateTime.Parse(Start);
            //DateTime end2 = string.IsNullOrEmpty(end) ? DateTime.MaxValue : DateTime.Parse(end);
            //TimeSpan ts = new TimeSpan(23, 59, 59);
            //end2 = end2.Date + ts;


            try
            {
                ViewBag.m = msg;
                var ClassInfo = new List<SelectListItem>
                {
                      new SelectListItem() { Text = "Select State", Selected = true, Value = "" }
                };
                ClassInfo.AddRange(db.taxes.Where(CCVV => CCVV.state != null).Select(b => new SelectListItem

                {
                    Text = b.state,
                    Selected = false,
                    Value = b.state.ToString()
                }).Distinct());
                ViewBag.lead_type = ClassInfo;



                var ClassInfo1 = new List<SelectListItem>
                {
                      new SelectListItem() { Text = "Select County", Selected = true, Value = "" }
                };
                ClassInfo1.AddRange(db.taxes.Where(CCVV => CCVV.county != null && CCVV.county != "" && CCVV.state == Type).Select(b => new SelectListItem

                {
                    Text = b.county,
                    Selected = true,
                    Value = b.county.ToString()
                }).Distinct());
                ViewBag.lead_type_county = ClassInfo1;
                //var result = db.taxes.Where(l => (l.state == Type) && (l.county == Type1 || l.county == "") && (l.tax_anme == "County Tax") && l.state == Type).ToList();

                if (Type == ""|| Type==null)
                {
                    var result = db.taxes
                        .Where(l => (l.tax_anme == "County Tax" && l.deleted == false))
                        .OrderBy(l => l.state).ThenBy(l => l.county).ThenBy(l => l.start_date)
                        .ToList();
                    CountyTaxController.lst = result;
                    return View(result.ToPagedList(page ?? 1, 8));
                }
                else
                { 
                    if (Type1 != "")
                    {
                        var  result = db.taxes
                        .Where(l => (l.tax_anme == "County Tax" && l.state == Type && l.deleted == false && l.county==Type1))
                        .OrderBy(l => l.state).ThenBy(l => l.county).ThenBy(l=> l.start_date)
                        .ToList();
                        CountyTaxController.lst = result;
                        return View(result.ToPagedList(page ?? 1, 8));
                    }
                    else
                    { 
                        var result = db.taxes
                            .Where(l => (l.tax_anme == "County Tax" && l.state == Type && l.deleted == false))
                            .OrderBy(l => l.state).ThenBy(l => l.county).ThenBy(l => l.start_date)
                            .ToList();
                        CountyTaxController.lst = result;
                        return View(result.ToPagedList(page ?? 1, 8));
                    }
                }

                //return View(db.leads.Where(x => (x.project_name.Contains(search) || search == null) && x.project_status_number == type && (x.project_status_number != 6 && x.deleted == false)).ToList());
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Add(FormCollection form)
        {
            try
            {

                tax target = new tax();

                //get property

                TryUpdateModel(target, new string[] { "tax_anme", "tax_value", "city", "state", "county", "zipcode", "start_date", "end_date", "deleted", "in_city" }, form.ToValueProvider());

                target.tax_anme = "County Tax";
                target.city = "";                
                target.zipcode = "";
                target.deleted = false;
                target.in_city = false;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;


                // Update Previous Tax Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = (DateTime)target.start_date;
                dt1 = dt1.AddDays(-1);

               
                if (db.taxes.Any(p => p.state == target.state) && db.taxes.Any(p => p.tax_anme == "County Tax") && db.taxes.Any(p=> p.county==target.county && p.deleted == false))
                {
                    ViewBag.m = "Duplicate Tax.";
                    return View();
                }
                else
                {
                    var ctx1 = new ckdatabase();
                    query = "Select Top 1 tax_number from taxes where tax_anme = 'County Tax' and state = '" + target.state + "' and county='" + target.county + "' and start_date<'" +
                        target.start_date + "' and deleted=0 order by start_date Desc";

                    PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                    if (PrevId != 0)
                    {
                        query = "Update taxes set end_date='" + dt1 + "' where tax_anme='County Tax' and state='" + target.state + "' and county='" + target.county + "'";
                        query = query + " and tax_number='" + PrevId.ToString() + "'";

                        var ctx2 = new ckdatabase();
                        int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                    }


                    db.taxes.Add(target);
                    db.SaveChanges();
                    ViewBag.m = "The County Tax was successfully created " + "on " + System.DateTime.Now;
                    return View(target);
                    string search = null;
                    return RedirectToAction("CountyTaxNew", new { search, msg = ViewBag.m });
                }
            }
            catch (Exception e)
            {
                ViewBag.m = "The County Tax was not created " + e.Message;
                return View();
            }
        }

        public ActionResult Add()
        {
            try
            {
                var utype = new List<SelectListItem>();
                utype.AddRange(db.users_types.Where(x => x.deleted != true).Select(a => new SelectListItem
                {
                    Text = a.user_type_name,
                    Selected = false,
                    Value = a.user_type_number.ToString()
                }));

                var branchtypes = new List<SelectListItem>();
                branchtypes.AddRange(db.branches.Where(x => x.deleted != true).Select(b => new SelectListItem
                {
                    Text = b.branch_name,
                    Selected = false,
                    Value = b.branch_number.ToString()
                }));

                ViewBag.utype = utype;
                ViewBag.branch = branchtypes;

                return View();
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }

        public ActionResult Delete(int id)
        {
            try
            {
                List<tax> CountyTax_list = db.taxes.Where(d => d.tax_number == id).ToList();
                ViewBag.Customerslist = CountyTax_list;
                tax target = CountyTax_list[0];
                target.deleted = true;
                db.SaveChanges();
                ViewBag.m = "The count Tax was successfully deleted.";
                string search = null;
                return RedirectToAction("CountyTaxNew", new { search, msg = ViewBag.m });
            }

            catch (Exception e)
            {
                ViewBag.m = "The County Tax was not deleted ..." + e.Message;
                string search = null;
                return RedirectToAction("CountyTaxNew", new { search, msg = ViewBag.m });
            }
        }


        public ActionResult Edit(int id)
        {
            try
            {
                //setting dropdown list for forgern key
                var utype = new List<SelectListItem>();
                utype.AddRange(db.users_types.Where(x => x.deleted != true).Select(a => new SelectListItem
                {
                    Text = a.user_type_name,
                    Selected = false,
                    Value = a.user_type_number.ToString()
                }));

                var branchtypes = new List<SelectListItem>();
                branchtypes.AddRange(db.branches.Where(x => x.deleted != true).Select(b => new SelectListItem
                {
                    Text = b.branch_name,
                    Selected = false,
                    Value = b.branch_number.ToString()
                }));
                //setting variable passing
                ViewBag.utype = utype;
                ViewBag.branch = branchtypes;
                ViewBag.id = id;


                List<tax> tax_list = db.taxes.Where(d => d.tax_number == id).ToList();
                ViewBag.Customerslist = tax_list;
                tax target = tax_list[0];

                //branchtypes.Where(q => int.Parse(q.Value) == target.branch.branch_number).First().Selected = true;
                //utype.Where(t => int.Parse(t.Value) == target.users_types.user_type_number).First().Selected = true;
                return View(target);
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }

        // Write to the DB that is why we use POST
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Edit(int id, FormCollection fo)
        {
            try
            {

                List<tax> tax_list = db.taxes.Where(d => d.tax_number == id).ToList();
                ViewBag.Employeeslist = tax_list;
                tax target = tax_list[0];

                TryUpdateModel(target, new string[] { "tax_anme", "tax_value", "city", "state", "county", "zipcode", "start_date", "end_date", "deleted", "in_city" }, fo.ToValueProvider());

                target.tax_anme = "County Tax";
                target.city = "";
                target.zipcode = "";
                target.deleted = false;
                target.in_city = false;
                ViewBag.Error = null;

                // Update Previous Tax Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = (DateTime)target.start_date;
                DateTime dt2 = new DateTime(2222, 1, 1); ;
                dt1 = dt1.AddDays(-1);

                var ctx1 = new ckdatabase();
                query = "Select Top 1 tax_number from taxes where tax_anme = 'County Tax' and state = '" + target.state + "' and start_date<'" + target.start_date +
                    "' and county='" + target.county + "' and deleted=0 and tax_number<> '" + id.ToString() + "' order by start_date Desc";
                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();

                if (PrevId != 0)
                {
                    query = "Update taxes set end_date='" + dt1 + "' where tax_anme='County Tax' and state='" + target.state + "' and county='" + target.county + "'";
                    query = query + " and tax_number='" + PrevId.ToString() + "'";

                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }


                db.SaveChanges();

                ViewBag.m = " The county tax was successfully updated " + " on " + System.DateTime.Now;
                //return View(target);

                string search = null;
                return RedirectToAction("CountyTaxNew", new { search, msg = ViewBag.m });
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }

        public ActionResult Update(int id)
        {
            try
            {
                //setting dropdown list for forgern key
                var utype = new List<SelectListItem>();
                utype.AddRange(db.users_types.Where(x => x.deleted != true).Select(a => new SelectListItem
                {
                    Text = a.user_type_name,
                    Selected = false,
                    Value = a.user_type_number.ToString()
                }));

                var branchtypes = new List<SelectListItem>();
                branchtypes.AddRange(db.branches.Where(x => x.deleted != true).Select(b => new SelectListItem
                {
                    Text = b.branch_name,
                    Selected = false,
                    Value = b.branch_number.ToString()
                }));
                //setting variable passing
                ViewBag.utype = utype;
                ViewBag.branch = branchtypes;
                ViewBag.id = id;


                List<tax> tax_list = db.taxes.Where(d => d.tax_number == id).ToList();
                ViewBag.Customerslist = tax_list;
                tax target = tax_list[0];

                //branchtypes.Where(q => int.Parse(q.Value) == target.branch.branch_number).First().Selected = true;
                //utype.Where(t => int.Parse(t.Value) == target.users_types.user_type_number).First().Selected = true;
                return View(target);
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Update(int id, FormCollection form)
        {
            try
            {

                tax target = new tax();

                //get property

                TryUpdateModel(target, new string[] { "tax_anme", "tax_value", "city", "state", "county", "zipcode", "start_date", "end_date", "deleted", "in_city" }, form.ToValueProvider());

                target.tax_anme = "County Tax";
                target.city = "";
                target.zipcode = "";
                target.deleted = false;
                target.in_city = false;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;


                // Update Previous Tax Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = (DateTime)target.start_date;
                dt1 = dt1.AddDays(-1);

                var ctx1 = new ckdatabase();
                query = "Select Top 1 tax_number from taxes where tax_anme = 'County Tax' and state = '" + target.state + "' and start_date<'" + target.start_date +
                    "' and county='" + target.county + "' and deleted=0 order by start_date Desc";

                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                if (PrevId != 0)
                {
                    query = "Update taxes set end_date='" + dt1 + "' where tax_anme='County Tax' and state='" + target.state + "' and county='" + target.county + "'";
                    query = query + " and tax_number='" + PrevId.ToString() + "'";

                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }


                db.taxes.Add(target);
                db.SaveChanges();
                ViewBag.m = "The County Tax was successfully created " + "on " + System.DateTime.Now;
                //return View(target);
                string search = null;
                return RedirectToAction("CountyTaxNew", new { search, msg = ViewBag.m });

            }
            catch (Exception e)
            {
                ViewBag.m = "The County Tax was not created " + e.Message;
                return View();
            }
        }
    }
}