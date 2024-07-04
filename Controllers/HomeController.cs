using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Web.Mvc;

namespace CreateWordSample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CreateDocument()
        {
            //Opens the Word template document
            using (WordDocument document = new WordDocument(ResolveApplicationDataPath1("/") + "/Letter Formatting.docx"))
            {
                string[] fieldNames = { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
				string[] fieldValues = { "Abhishek Alandikar", "Syncfusion", "507 - 20th Ave. E.Apt. 2A", "Seattle, WA", "USA", "(206) 555-9857-x5467" };
				
				//Performs the mail merge
				document.MailMerge.Execute(fieldNames, fieldValues);
                
                //Saves the Word document to disk in DOCX format
                document.Save("Result.docx", FormatType.Docx, HttpContext.ApplicationInstance.Response, HttpContentDisposition.Attachment);
            }

            return View();
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
		
		protected string ResolveApplicationDataPath1(string fileName)
        {
            //Data folder path is resolved from requested page physical path.
            string dataPath = new System.IO.DirectoryInfo(Request.PhysicalPath + "..\\..\\..\\Data\\").FullName;
            return string.Format("{0}\\{1}", dataPath, fileName);
        }
    }
}