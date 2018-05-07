
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using PagedList;
using PagedList.Mvc;

namespace ck_project.Controllers
{
    [Authorize(Roles = "Administrator")]
    public class StateTaxController : Controller
    {
        //Creating the db connecton
        ckdatabase db = new ckdatabase();
        static List<tax> lst = new List<tax>();
               
        public ActionResult ListTax1(int? page, string search, String msg = null)
        {
            try
            {
                ViewBag.m = msg;
                if (search == "")
                {
                    search = null;
                }
                
                return View(db.taxes
                    .Where(x => x.state.Contains(search) || search == null && x.deleted==false && x.tax_anme=="State Tax")
                    .OrderBy(l => l.state).ThenBy(l=> l.start_date)
                    .ToList().ToPagedList(page ?? 1, 8));
            }
            catch (Exception e)
            {
                ViewBag.m = e.Message;
                return View();
            }
        }

        // for Dropdown List
        public ActionResult ListTax(int? page, string search = null, string Type = null, string Type1 = null, string Start = null, string end = null, string msg = null)
        {


            DateTime start = string.IsNullOrEmpty(Start) ? DateTime.MinValue : DateTime.Parse(Start);
            DateTime end2 = string.IsNullOrEmpty(end) ? DateTime.MaxValue : DateTime.Parse(end);
            TimeSpan ts = new TimeSpan(23, 59, 59);
            end2 = end2.Date + ts;


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
                
                if (search == "")
                {
                    search = null;
                }

                return View(db.taxes
                    .Where(x => x.tax_anme.Trim() == "State Tax" && (x.state.Contains(search) || search == null) && x.deleted == false)
                    .OrderBy(l => l.state).ThenBy(l => l.start_date)
                    .ToList().ToPagedList(page ?? 1, 8));
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

                target.tax_anme = "State Tax";
                target.city= "";
                target.county= "";
                target.zipcode = "";
                target.deleted= false;
                target.in_city= false ;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;


                // Update Previous Tax Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1= (DateTime)target.start_date;
                dt1 = dt1.AddDays(-1);
                

                if (db.taxes.Any(p => p.state == target.state && p.tax_anme == "State Tax" && p.deleted == false))
                { 
                        ViewBag.m = "Duplicate State.";
                        return View();
                }
                else
                {
                    var ctx1 = new ckdatabase();
                    query = "Select Top 1 tax_number from taxes where tax_anme = 'State Tax' and state = '" + target.state + "' and start_date<'" + target.start_date +
                        "' and deleted=0 order by start_date Desc";

                    PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                    if (PrevId != 0)
                    {
                        query = "Update taxes set end_date='" + dt1 + "' where tax_anme='State Tax' and state='" + target.state + "'";
                        query = query + " and tax_number='" + PrevId.ToString() + "'";

                        var ctx2 = new ckdatabase();
                        int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                    }


                    db.taxes.Add(target);
                    db.SaveChanges();
                    ViewBag.m = "The State Tax was successfully created " + "on " + System.DateTime.Now;
                    string search = null;
                    return RedirectToAction("ListTax", new { search, msg = ViewBag.m });
                }
            }
            catch (Exception e)
            {
                ViewBag.m = "The State Tax was not created " + e.Message;
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
                List<tax> StateTax_list = db.taxes.Where(d => d.tax_number== id).ToList();
                ViewBag.Customerslist = StateTax_list;
                tax target = StateTax_list[0];
                target.deleted = true;
                db.SaveChanges();
                ViewBag.m = "The State Tax was successfully deleted.";
                string search = null;
                return RedirectToAction("ListTax", new { search , msg = ViewBag.m });
            }


            catch (Exception e)
            {
                ViewBag.m = "The State Tax was not deleted ..." + e.Message;
                string search = null;
                return RedirectToAction("ListTax", new { search, msg = ViewBag.m });
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


                List<tax> tax_list = db.taxes.Where(d => d.tax_number== id).ToList();
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
                
                List<tax> tax_list = db.taxes.Where(d => d.tax_number== id).ToList();
                ViewBag.Employeeslist = tax_list;
                tax target = tax_list[0];

                TryUpdateModel(target, new string[] { "tax_anme", "tax_value", "city", "state", "county", "zipcode", "start_date", "end_date", "deleted", "in_city" }, fo.ToValueProvider());

                target.tax_anme = "State Tax";
                target.city = "";
                target.county = "";
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
                query = "Select Top 1 tax_number from taxes where tax_anme = 'State Tax' and state = '" + target.state + "' and start_date<'" + target.start_date +
                    "' and deleted=0 and tax_number<> '" + id.ToString() + "' order by start_date Desc";
                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();

                if (PrevId != 0)
                {
                    query = "Update taxes set end_date='" + dt1 + "' where tax_anme='State Tax' and state='" + target.state + "'";
                    query = query + " and tax_number='" + PrevId.ToString() + "'";

                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }


                db.SaveChanges();

                ViewBag.m = " The state tax was successfully updated " + " on " + System.DateTime.Now;
                //return View(target);

                string search = null;
                return RedirectToAction("ListTax", new { search, msg = ViewBag.m });
            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
                return View();
            }
        }

        public ActionResult Update (int id)
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

                target.tax_anme = "State Tax";
                target.city = "";
                target.county = "";
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
                query = "Select Top 1 tax_number from taxes where tax_anme = 'State Tax' and state = '" + target.state + "' and start_date<'" + target.start_date +
                    "' and deleted=0 order by start_date Desc";

                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                if (PrevId != 0)
                {
                    query = "Update taxes set end_date='" + dt1 + "' where tax_anme='State Tax' and state='" + target.state + "'";
                    query = query + " and tax_number='" + PrevId.ToString() + "'";

                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }


                db.taxes.Add(target);
                db.SaveChanges();
                ViewBag.m = "The State Tax was successfully created " + "on " + System.DateTime.Now;
                //return View(target);
                string search = null;
                return RedirectToAction("ListTax", new { search, msg = ViewBag.m });
                
            }
            catch (Exception e)
            {
                ViewBag.m = "The State Tax was not created " + e.Message;
                return View();
            }
        }
    }
}