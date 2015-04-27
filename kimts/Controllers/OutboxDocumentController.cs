using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using DataModel;

namespace kimts.Controllers
{
    public class OutboxDocumentController : Controller
    {
        private okimtsDbEntities db = new okimtsDbEntities();

        // GET: /OutboxDocument/
        public ActionResult Index()
        {
            var outboxdocuments = db.OutboxDocuments.Include(o => o.BuildinObject).Include(o => o.ContractorEmploye).Include(o => o.Contractor).Include(o => o.DocState1).Include(o => o.Employe).Include(o => o.Employe1).Include(o => o.FileMetaData).Include(o => o.InboxDocument).Include(o => o.TypesOfOutboxDoc);
            return View(outboxdocuments.ToList());
        }

        // GET: /OutboxDocument/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            OutboxDocument outboxdocument = db.OutboxDocuments.Find(id);
            if (outboxdocument == null)
            {
                return HttpNotFound();
            }
            return View(outboxdocument);
        }

        // GET: /OutboxDocument/Create
        public ActionResult Create()
        {
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName");
            ViewBag.RecieverEmploye = new SelectList(db.ContractorEmployes, "Id", "FullName");
            ViewBag.RecieverOrg = new SelectList(db.Contractors, "Id", "OrgName");
            ViewBag.DocState = new SelectList(db.DocStates, "Id", "StateName");
            ViewBag.WhoMade = new SelectList(db.Employes, "Id", "FullName");
            ViewBag.WhoSign = new SelectList(db.Employes, "Id", "FullName");
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName");
            ViewBag.ResponseOn = new SelectList(db.InboxDocuments, "Id", "SenderNum");
            ViewBag.TypeOfOutboxDoc = new SelectList(db.TypesOfOutboxDocs, "Id", "NameOfType");
            return View();
        }

        // POST: /OutboxDocument/Create
        // Чтобы защититься от атак чрезмерной передачи данных, включите определенные свойства, для которых следует установить привязку. Дополнительные 
        // сведения см. в статье http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include="Id,RecieverOrg,RecieverEmploye,BuildingObj,OutboxNum,OutboxDate,DocTheme,WhoSign,WhoMade,ResponseOn,SentDate,DocState,Files,TypeOfOutboxDoc")] OutboxDocument outboxdocument)
        {
            if (ModelState.IsValid)
            {
                db.OutboxDocuments.Add(outboxdocument);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", outboxdocument.BuildingObj);
            ViewBag.RecieverEmploye = new SelectList(db.ContractorEmployes, "Id", "FullName", outboxdocument.RecieverEmploye);
            ViewBag.RecieverOrg = new SelectList(db.Contractors, "Id", "OrgName", outboxdocument.RecieverOrg);
            ViewBag.DocState = new SelectList(db.DocStates, "Id", "StateName", outboxdocument.DocState);
            ViewBag.WhoMade = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoMade);
            ViewBag.WhoSign = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoSign);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", outboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.InboxDocuments, "Id", "SenderNum", outboxdocument.ResponseOn);
            ViewBag.TypeOfOutboxDoc = new SelectList(db.TypesOfOutboxDocs, "Id", "NameOfType", outboxdocument.TypeOfOutboxDoc);
            return View(outboxdocument);
        }

        // GET: /OutboxDocument/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            OutboxDocument outboxdocument = db.OutboxDocuments.Find(id);
            if (outboxdocument == null)
            {
                return HttpNotFound();
            }
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", outboxdocument.BuildingObj);
            ViewBag.RecieverEmploye = new SelectList(db.ContractorEmployes, "Id", "FullName", outboxdocument.RecieverEmploye);
            ViewBag.RecieverOrg = new SelectList(db.Contractors, "Id", "OrgName", outboxdocument.RecieverOrg);
            ViewBag.DocState = new SelectList(db.DocStates, "Id", "StateName", outboxdocument.DocState);
            ViewBag.WhoMade = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoMade);
            ViewBag.WhoSign = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoSign);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", outboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.InboxDocuments, "Id", "SenderNum", outboxdocument.ResponseOn);
            ViewBag.TypeOfOutboxDoc = new SelectList(db.TypesOfOutboxDocs, "Id", "NameOfType", outboxdocument.TypeOfOutboxDoc);
            return View(outboxdocument);
        }

        // POST: /OutboxDocument/Edit/5
        // Чтобы защититься от атак чрезмерной передачи данных, включите определенные свойства, для которых следует установить привязку. Дополнительные 
        // сведения см. в статье http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include="Id,RecieverOrg,RecieverEmploye,BuildingObj,OutboxNum,OutboxDate,DocTheme,WhoSign,WhoMade,ResponseOn,SentDate,DocState,Files,TypeOfOutboxDoc")] OutboxDocument outboxdocument)
        {
            if (ModelState.IsValid)
            {
                db.Entry(outboxdocument).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", outboxdocument.BuildingObj);
            ViewBag.RecieverEmploye = new SelectList(db.ContractorEmployes, "Id", "FullName", outboxdocument.RecieverEmploye);
            ViewBag.RecieverOrg = new SelectList(db.Contractors, "Id", "OrgName", outboxdocument.RecieverOrg);
            ViewBag.DocState = new SelectList(db.DocStates, "Id", "StateName", outboxdocument.DocState);
            ViewBag.WhoMade = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoMade);
            ViewBag.WhoSign = new SelectList(db.Employes, "Id", "FullName", outboxdocument.WhoSign);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", outboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.InboxDocuments, "Id", "SenderNum", outboxdocument.ResponseOn);
            ViewBag.TypeOfOutboxDoc = new SelectList(db.TypesOfOutboxDocs, "Id", "NameOfType", outboxdocument.TypeOfOutboxDoc);
            return View(outboxdocument);
        }

        // GET: /OutboxDocument/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            OutboxDocument outboxdocument = db.OutboxDocuments.Find(id);
            if (outboxdocument == null)
            {
                return HttpNotFound();
            }
            return View(outboxdocument);
        }

        // POST: /OutboxDocument/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            OutboxDocument outboxdocument = db.OutboxDocuments.Find(id);
            db.OutboxDocuments.Remove(outboxdocument);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult Word(int? id)
        {
            OutboxDocument outboxdocument = db.OutboxDocuments.Find(id);
            Extensions.OutboxDocTamplate tmpl = new Extensions.OutboxDocTamplate(outboxdocument);

            //TODO: Generate filename realisation
            // - path to save

            // return message about whats to do with file: 
            // Открыть для редактирования или сохранить? Открыть, Сохранить, Отмена

            Extensions.MsWord.SearchAndReplace(@"C:\" + tmpl.GetFileName(), tmpl);

            return RedirectToAction("Index");
        }

        public string ReturnRegNumber ()
        {
            int lastNum = db.OutboxDocuments.Max(x => x.IndexNumber).Value;
            int IncIndexNum = lastNum++;

            return IncIndexNum.ToString();
        }
    }
}
