using System.Linq;
using System.Web.Mvc;
using CreateWordSample.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace CreateWordSample.Controllers
{
    public class HomeController : Controller
    {
        private DemoEntities db = new DemoEntities();

        // GET: Home/Index
        public ActionResult Index()
        {
            return View();
        }

        // GET: Home/CreateDocument
        public ActionResult CreateDocument()
        {
            return View();
        }

        // POST: Home/CreateDocument
        [HttpPost]
        public ActionResult CreateDocument(Customer model)
        {
            if (ModelState.IsValid)
            {
                // Map Customer model to Demo entity
                var demo = new Demo
                {
                    ContactName = model.ContactName,
                    CompanyName = model.CompanyName,
                    Address = model.Address,
                    City = model.City,
                    Country = model.Country,
                    Phone = model.Phone
                };

                // Add the entity to the DbSet and save changes
                db.Demoes.Add(demo);
                db.SaveChanges();

                // Opens the Word template document
                using (WordDocument document = new WordDocument(ResolveApplicationDataPath("SampleDocument.docx")))
                {
                    string[] fieldNames = { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
                    string[] fieldValues = { model.ContactName, model.CompanyName, model.Address, model.City, model.Country, model.Phone };

                    // Performs the mail merge
                    document.MailMerge.Execute(fieldNames, fieldValues);

                    // Saves the Word document to disk in DOCX format
                    document.Save("Result.docx", FormatType.Docx, HttpContext.ApplicationInstance.Response, HttpContentDisposition.Attachment);
                }

                // Redirect to a suitable action after processing
                return RedirectToAction("Index");
            }

            return View(model);
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

        protected string ResolveApplicationDataPath(string fileName)
        {
            // Get the path to the Data folder
            string dataPath = Server.MapPath("~/Data");
            return System.IO.Path.Combine(dataPath, fileName);
        }
    }
}
