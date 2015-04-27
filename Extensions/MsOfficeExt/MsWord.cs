using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
//using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Extensions.MsOfficeExt.TemplateForGenerate;
using DataModel;


namespace Extensions
{
    sealed  class Person
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string SecondName { get; set; }
    }

    public class OutboxDocTamplate
    {

        // public OutboxDocTamplate()
        //{
        //    this.OutboxDict = new Dictionary<int, string>()
        //        {
        //            {1,"[RecieverPost]"},
        //            {2,"[RecieverInitials]"},
        //            {3,"[OutboxTheme]"},
        //            {4,"[RecieverOrg]"},
        //            {5, "[WhoSignPost]"},
        //            {6, "[WhoSignName]"},
        //            {7,"[WhoMadeName]"},
        //            {8,"[WhoMadeTel]"},
        //            {9,"[DearReciever]!"},
        //            {10,"[OutboxDate]"},
        //            {11,"[OutboxNum]"}
        //        }; 

        // }
        private string recieverinitials;
        private string dearreciever;

        public OutboxDocTamplate( OutboxDocument odoc)
        {
            RecieverPost = odoc.ContractorEmploye.PostDateln;
            RecieverOrg = odoc.Contractor.OrgName;
            OutboxDate = odoc.OutboxDate.ToString();
            OutboxNum = odoc.OutboxNum;
            OutboxTheme = odoc.DocTheme;
            Gender = odoc.ContractorEmploye.Gender.Trim();
            FName = odoc.ContractorEmploye.FName.Trim();
            SName = odoc.ContractorEmploye.SName.Trim();
            LName = odoc.ContractorEmploye.LName.Trim();
            LNameDateln = odoc.ContractorEmploye.LNameDateln.Trim();
            RecieverInitials = "";
            DearReciever = "";
            WhoMadeName = GetInitials(odoc.Employe.FullName);
            WhoMadeTel = odoc.Employe.Phone;
            WhoSignName = GetInitials(odoc.Employe1.FullName);
            WhoSignPost = odoc.Employe1.Post;
        }

        public Dictionary<int, string> OutboxDict { get; set; }
        public string RecieverPost { get; set; } //"Начальнику" Прикрутить дательный падеж
        public string RecieverInitials 
        {
            get{return recieverinitials;} 
            private set {recieverinitials = FName.Substring(0, 1) + "." +
                SName.Substring(0, 1) + "."+" "+
                    LNameDateln.Trim();
            }
        }
        public string OutboxTheme { get; private set; }
        public string RecieverOrg { get; private set; }
        public string WhoSignPost { get; private set; }
        public string WhoSignName { get; private set; }
        public string WhoMadeName { get; private set; }
        public string WhoMadeTel { get; private set; }
        public string DearReciever
        { get { return dearreciever; } 
          private set 
          { 
              string end =  Gender.ToLower () == "муж" ? "ый" : "ая"  ;
              dearreciever = "Уважаем" + end +" "+ FName + " "+ SName+"!";
          }
        }
        public string OutboxDate { get; private set; }
        public string OutboxNum { get; private set; }
        public string FName { get; private set; }
        public string SName { get; private set; }
        public string LName { get; private set; }
        public string LNameDateln { get; private set; }
        public string Gender { get; private set; }

        public string GetFileName()
        {
            string nonspace = Regex.Replace(OutboxNum + " от  " + OutboxDate + " " + OutboxTheme + ".docx", "\\s+", " ");
            return Regex.Replace(nonspace, "(?:[^a-zA-Z_0-9.]|(?<=['\']))", "_");
        }

        private string GetInitials(string FullName)
        {
            string nonspace = Regex.Replace(FullName, "\\s+", " ");
            var spl = FullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (spl.Length != 3)
                throw new Exception (" ФИО должно быть задано следующим форматом: Фамилия Имя Отчество");
            Person prsn = new Person { LastName = spl[0], FirstName = spl[1], SecondName = spl[2] };
            string prsnStr = prsn.FirstName.Substring(0, 1) + "." +
                prsn.SecondName.Substring(0, 1) + ". " +
                    prsn.LastName.Trim();
            return prsnStr;
        }

    }


    public class MsWord
    {

        public static void SearchAndReplace(string document, OutboxDocTamplate tmpl)
        {
         
           //Change generate template to _PromzonaTemplate
           Extensions.MsOfficeExt.TemplateForGenerate._PromzonaTemplate pt = new Extensions.MsOfficeExt.TemplateForGenerate._PromzonaTemplate();
           pt.CreatePackage(document, tmpl);

        }
        }

}

