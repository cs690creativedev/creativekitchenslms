using PagedList;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ck_project.Controllers
{
    public class LaborCostController : Controller
    {
        // GET: LaborCost
        //Creating the db connecton
        ckdatabase db = new ckdatabase();
        static List<rate> lst = new List<rate>();
        public ActionResult ListLaborCost(int? page, string search = null, string Type = null, string Start = null, string end = null, string msg = null)
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
                    new SelectListItem() { Text = "Select Type", Selected = true, Value = "" }
                };
                ClassInfo.AddRange(db.rates.Where(CCVV => CCVV.rate_name != null).Select(b => new SelectListItem

                {
                    Text = b.rate_name,
                    Selected = false,
                    Value = b.rate_name.ToString()
                }).Distinct());
                ViewBag.lead_type = ClassInfo;

                if (Type == null|| Type == "Select Type" || Type == "")
                {
                    var result = db.rates
                        .Where(l => ((l.deleted == false)))
                        .OrderBy(l=> l.rate_name)
                        .ToList();
                    //LaborCostController.lst = result;
                    return View(result.ToPagedList(page ?? 1, 8));
                }
                else
                {
                    var result = db.rates
                        .Where(l => ((l.rate_name == Type && l.deleted == false)))
                        .OrderBy(l => l.rate_name)
                        .ToList();
                    //LaborCostController.lst = result;
                    return View(result.ToPagedList(page ?? 1, 8));
                }
                

            }
            catch (Exception e)
            {
                ViewBag.m = "Something went wrong ... " + e.Message;
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

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Add(FormCollection form)
        {
            try
            {

                rate target = new rate();

                //get property

                TryUpdateModel(target, new string[] { "rate_name", "rate_measurement", "amount", "start_date", "end_date", "deleted" }, form.ToValueProvider());

                target.deleted = false;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;

                // Update Previous Cost Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = target.start_date;
                dt1 = dt1.AddDays(-1);


                if (db.rates.Any(p => p.rate_name == target.rate_name && p.deleted==false  ))
                {
                    ViewBag.m = "Duplicate Entry.";
                    return View();
                }
                else
                {
                    var ctx1 = new ckdatabase();
                    query = "Select Top 1 Rate_number from rates where rate_name ='" + target.rate_name + "' and start_date<'" + target.start_date +
                        "' and deleted=0 order by start_date Desc";

                    PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                    if (PrevId != 0)
                    {
                        query = "Update rates set end_date='" + dt1 + "' where tax_number='" + PrevId.ToString() + "'";
                        var ctx2 = new ckdatabase();
                        int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                    }


                    db.rates.Add(target);
                    db.SaveChanges();
                    ViewBag.m = "The Labor Cost was successfully created " + "on " + System.DateTime.Now;
                    string search = null;
                    return RedirectToAction("ListLaborCost", new { search, msg = ViewBag.m });
                }
            }
            catch (Exception e)
            {
                ViewBag.m = "The Labor was not created " + e.Message;
                return View();
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


                List<rate> rate_list = db.rates.Where(d => d.rate_number == id).ToList();
                ViewBag.Customerslist = rate_list;
                rate target = rate_list[0];

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
        public ActionResult Edit(int id, FormCollection form)
        {
            try
            {

                List<rate> rate_list = db.rates.Where(d => d.rate_number == id).ToList();
                ViewBag.Employeeslist = rate_list;
                rate target = rate_list[0];

                TryUpdateModel(target, new string[] { "rate_name", "rate_measurement", "amount", "start_date", "end_date", "deleted" }, form.ToValueProvider());

                target.deleted = false;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;

                // Update Previous Labor Cost Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = target.start_date;
                DateTime dt2 = new DateTime(2222, 1, 1); ;
                dt1 = dt1.AddDays(-1);

                var ctx1 = new ckdatabase();
                query = "Select Top 1 rate_number from rates where deleted=0 and rate_number<> '" + id.ToString() + "' order by start_date Desc";
                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();

                if (PrevId != 0)
                {
                    query = "Update rates set end_date='" + dt1 + "' where rate_number='" + PrevId.ToString() + "'";

                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }
                
                db.SaveChanges();
                ViewBag.m = " The Labor Cost was successfully updated " + " on " + System.DateTime.Now;
                //return View(target);

                string search = null;
                return RedirectToAction("ListLaborCost", new { search, msg = ViewBag.m });
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


                List<rate> rate_list = db.rates.Where(d => d.rate_number == id).ToList();
                ViewBag.Customerslist = rate_list;
                rate target = rate_list[0];

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

                rate target = new rate();

                //get property

                TryUpdateModel(target, new string[] { "rate_name", "rate_measurement", "amount",  "start_date", "end_date", "deleted"}, form.ToValueProvider());
                
                target.deleted = false;
                target.end_date = new DateTime(2222, 1, 1);
                ViewBag.Error = null;


                // Update Previous Tax Detail
                string query = "";
                int PrevId = 0;
                DateTime dt1 = target.start_date;
                dt1 = dt1.AddDays(-1);

                var ctx1 = new ckdatabase();
                query = "Select Top 1 rate_number from rates where rate_name ='" + target.rate_name + "' and start_date<'" + target.start_date +
                    "' and deleted=0 order by start_date Desc";

                PrevId = ctx1.Database.SqlQuery<int>(query).FirstOrDefault<int>();
                if (PrevId != 0)
                {
                    query = "Update rates set end_date='" + dt1 + "' where rate_number='" + PrevId.ToString() + "'";
                    var ctx2 = new ckdatabase();
                    int noOfRowUpdated = ctx2.Database.ExecuteSqlCommand(query);
                }


                db.rates.Add(target);
                db.SaveChanges();
                ViewBag.m = "The Labor cost was successfully created " + "on " + System.DateTime.Now;
                //return View(target);
                string search = null;
                return RedirectToAction("ListLaborCost", new { search, msg = ViewBag.m });

            }
            catch (Exception e)
            {
                ViewBag.m = "The Labor Cost was not created " + e.Message;
                return View();
            }
        }

        public ActionResult Delete(int id)
        {
            try
            {
                List<rate> rate_list = db.rates.Where(d => d.rate_number == id).ToList();
                ViewBag.Customerslist = rate_list;
                rate target = rate_list[0];
                target.deleted = true;
                db.SaveChanges();
                ViewBag.m = "The Labor Cost was successfully deleted.";
                string search = null;
                return RedirectToAction("ListLaborCost", new { search, msg = ViewBag.m });
            }


            catch (Exception e)
            {
                ViewBag.m = "The Labor Cost was not deleted ..." + e.Message;
                string search = null;
                return RedirectToAction("ListLaborCost", new { search, msg = ViewBag.m });
            }
        }

    }
}