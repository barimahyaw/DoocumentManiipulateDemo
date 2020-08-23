using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using System;
using System.IO;

namespace DoocumentManiipulateDemo.Controllers
{
    /// <summary>
    /// Easily Manipulate document using Spire.Doc ( generate, open, write, edit and save document)
    /// </summary>
    public class DocumentController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public DocumentController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        /// <summary>
        ///  edit Word document, convert to pdf and render in browser without saving pdf to local drive
        /// </summary>
        /// <returns>pdf using FileStreamResult</returns>
        public IActionResult SampleTest()
        {
            var guarantor = " ";

            // Get Path of Template of Word/Pdf/Excel/Image/Html file from your local drive
            var path = Path.Combine(_hostingEnvironment.WebRootPath, "Templates/GuarantorReleasedCertificate.docx");


            using var document = new Document(path);

            // Perform the replacement of placeholders:    
            var guarantorName = !string.IsNullOrWhiteSpace(guarantor)
                ? "Guarantor Institution Name"
                : "Esther Eshun";
            const string borrowerName = "Asamoah, Joseph Kwame";

            document.Replace("%DatePrepared%", DateTime.Now.ToString("D"), false, true);
            document.Replace("%GuarantorName%", guarantorName, false, true);
            document.Replace("%GuarantorSsf%", "E1045739362673", false, true);
            document.Replace("%GuarantorPostalAddress%", "Ghana Health Service", false, true);
            document.Replace("%BorrowerName%", borrowerName, false, true);
            document.Replace("%BorrowerSsf%", "E206453434288", false, true);
            document.Replace("%BorrowerPostalAddress%", "P. O. Box 5, Koforidua", false, true);
            document.Replace("%Institution%", "Kwame Nkrumah University of Science and Technology", false, true);
            document.Replace("%GuarantorReleasedBy%", "Eric Mawudeku", false, true);
            document.Replace("%Destination%", "Administrator", false, true);
            document.Replace("%OfficeBranch%", "Accra Main", false, true);
            document.Replace("%OfficeBranch%", "Accra Main", false, true);

            //access memory stream
            var ms = new MemoryStream();

            // Save the document to memory
            document.SaveToStream(ms, FileFormat.PDF);

            //specify the position of file in the stream to use for seek
            ms.Seek(0, SeekOrigin.Begin);

            //prepare unique name for generated pdf and save to local drive
            /*var fileName = $"GuarantorReleaseCertificate_E19098645_{DateTime.Now:ddMMyyyyHHmmss}";
            var htmlFilePath = Path.Combine(_hostingEnvironment.WebRootPath, $"Templates/{fileName}.pdf");
            document.SaveToFile(htmlFilePath, FileFormat.PDF);*/

            //Open and read saved pdf document in memoryStream
            //var fs = new FileStream(htmlFilePath, FileMode.Open, FileAccess.Read);

            //Delete the saved PDF File.
            //System.IO.File.Delete(htmlFilePath);

            //return generated pdf from stream to display in browser
            var fileResult = File(ms, "application/pdf");
            return fileResult;
        }

    }
}
