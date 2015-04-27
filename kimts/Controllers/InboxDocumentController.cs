using DataModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;



namespace kimts.Controllers
{
    public class InboxDocumentController : Controller
    {
        private  okimtsDbEntities db = new okimtsDbEntities();

        // GET: /InboxDocument/
        public ActionResult Index()
        {
            var inboxdocuments = db.InboxDocuments.Include(i => i.BuildinObject).Include(i => i.ContractorEmploye).Include(i => i.Employe).Include(i => i.FileMetaData).Include(i => i.OutboxDocument);
            return View(inboxdocuments.ToList());
        }

        // GET: /InboxDocument/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            InboxDocument inboxdocument = db.InboxDocuments.Find(id);
            if (inboxdocument == null)
            {
                return HttpNotFound();
            }
            return View(inboxdocument);
        }

        // GET: /InboxDocument/Create
        public ActionResult Create()
        {
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName");
            ViewBag.Sender = new SelectList(db.ContractorEmployes, "Id", "FullName");
            ViewBag.Reciever = new SelectList(db.Employes, "Id", "FullName");
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName");
            ViewBag.ResponseOn = new SelectList(db.OutboxDocuments, "Id", "OutboxNum");
            return View();
        }

        // POST: /InboxDocument/Create
        // Чтобы защититься от атак чрезмерной передачи данных, включите определенные свойства, для которых следует установить привязку. Дополнительные 
        // сведения см. в статье http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include="Id,Sender,SenderNum,SenderDate,ResponseOn,Reciever,DocTheme,DocState,BuildingObj,Files")] InboxDocument inboxdocument)
        {
            if (ModelState.IsValid)
            {
                db.InboxDocuments.Add(inboxdocument);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", inboxdocument.BuildingObj);
            ViewBag.Sender = new SelectList(db.ContractorEmployes, "Id", "FullName", inboxdocument.Sender);
            ViewBag.Reciever = new SelectList(db.Employes, "Id", "FullName", inboxdocument.Reciever);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", inboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.OutboxDocuments, "Id", "OutboxNum", inboxdocument.ResponseOn);
            return View(inboxdocument);
        }

        // GET: /InboxDocument/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            InboxDocument inboxdocument = db.InboxDocuments.Find(id);
            if (inboxdocument == null)
            {
                return HttpNotFound();
            }
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", inboxdocument.BuildingObj);
            ViewBag.Sender = new SelectList(db.ContractorEmployes, "Id", "FullName", inboxdocument.Sender);
            ViewBag.Reciever = new SelectList(db.Employes, "Id", "FullName", inboxdocument.Reciever);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", inboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.OutboxDocuments, "Id", "OutboxNum", inboxdocument.ResponseOn);
            return View(inboxdocument);
        }

        // POST: /InboxDocument/Edit/5
        // Чтобы защититься от атак чрезмерной передачи данных, включите определенные свойства, для которых следует установить привязку. Дополнительные 
        // сведения см. в статье http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include="Id,Sender,SenderNum,SenderDate,ResponseOn,Reciever,DocTheme,DocState,BuildingObj,Files")] InboxDocument inboxdocument)
        {
            if (ModelState.IsValid)
            {
                db.Entry(inboxdocument).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.BuildingObj = new SelectList(db.BuildinObjects, "Id", "ObjName", inboxdocument.BuildingObj);
            ViewBag.Sender = new SelectList(db.ContractorEmployes, "Id", "FullName", inboxdocument.Sender);
            ViewBag.Reciever = new SelectList(db.Employes, "Id", "FullName", inboxdocument.Reciever);
            ViewBag.Files = new SelectList(db.FileMetaDatas, "Id", "FileName", inboxdocument.Files);
            ViewBag.ResponseOn = new SelectList(db.OutboxDocuments, "Id", "OutboxNum", inboxdocument.ResponseOn);
            return View(inboxdocument);
        }

        // GET: /InboxDocument/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            InboxDocument inboxdocument = db.InboxDocuments.Find(id);
            if (inboxdocument == null)
            {
                return HttpNotFound();
            }
            return View(inboxdocument);
        }

        // POST: /InboxDocument/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            InboxDocument inboxdocument = db.InboxDocuments.Find(id);
            db.InboxDocuments.Remove(inboxdocument);
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
    }
}
