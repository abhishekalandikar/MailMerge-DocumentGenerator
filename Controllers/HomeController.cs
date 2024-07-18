using System.Linq;
using System.Web.Mvc;
using CreateWordSample.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace CreateWordSample.Controllers
{
    public class HomeController : Controller
    {
        private CompanyEntities db = new CompanyEntities();

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
                var demo = new Company
                {
                    CustomerId = model.CustomerId,
                    Name = model.Name,
                    Address = model.Address,
                    PhoneNumber = model.PhoneNumber,
                    Aadhar = model.Aadhar,
                    PAN = model.PAN
                };

                // Add the entity to the DbSet and save changes
                db.Companies.Add(demo);
                db.SaveChanges();

                // Opens the Word template document
                using (WordDocument document = new WordDocument(ResolveApplicationDataPath("SampleDocument.docx")))
                {
                    string[] fieldNames = { "CustomerId", "Name", "Address", "PhoneNumber", "Aadhar", "PAN" };
                    string[] fieldValues = { model.CustomerId.ToString(), model.Name, model.Address, model.PhoneNumber, model.Aadhar, model.PAN };

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

        

        

        protected string ResolveApplicationDataPath(string fileName)
        {
            // Get the path to the Data folder
            string dataPath = Server.MapPath("~/Data");
            return System.IO.Path.Combine(dataPath, fileName);
        }
    }
}
