using DataModel;
using Extensions.Models.jsTree;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Extensions.Controllers
{
    public class CompanyStructureController : Controller
    {
        private okimtsDbEntities db = new okimtsDbEntities();
        //
        // GET: /CompanyStructure/
        public ActionResult Index()
        {
            List<CompanyStructure> allNodes = new List<CompanyStructure>();
            allNodes = db.CompanyStructures.ToList();
            var rootNode = allNodes.Where(x => x.ParentDepartmentId == null).FirstOrDefault();
            SetChildren(rootNode, allNodes);
            return View(rootNode);
        }

        private void SetChildren(CompanyStructure model, List<CompanyStructure> depList)
        {
            var childs = depList.Where(x => x.ParentDepartmentId == model.Id).ToList();
            if (childs.Count > 0)
            {
                foreach (var child in childs)
                {
                    SetChildren(child, depList);
                    model.ChildNodes.Add(child);
                }
            }
        }
    }
}