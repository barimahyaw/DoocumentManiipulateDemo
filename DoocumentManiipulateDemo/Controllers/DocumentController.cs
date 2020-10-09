using Microsoft.AspNetCore.Mvc;

namespace DoocumentManiipulateDemo.Controllers
{
    /// <summary>
    /// Easily Manipulate document using Spire.Doc ( generate, open, write, edit and save document)
    /// </summary>
    public class DocumentController : Controller
    {
        private readonly IDocumentMerger _documentMerger;

        public DocumentController(IDocumentMerger documentMerger)
        {
            _documentMerger = documentMerger;
        }
        /// <summary>
        ///  edit Word document, convert to pdf and render in browser without saving pdf to local drive
        /// </summary>
        /// <returns>pdf using FileStreamResult</returns>
        public IActionResult SampleTest()
        {
            //return generated pdf from stream to display in browser
            return File(_documentMerger.TestDocument(), "application/pdf");
        }


        /// <summary>
        /// edit two Word document, merge and convert to pdf and render in browser without saving pdf to local drive
        /// </summary>
        /// <returns></returns>
        public IActionResult SampleMergedTest()
        {
            //return generated merged pdf from stream to display in browser
            return File(_documentMerger.PerformDocumentMerge(), "application/pdf");
        }
    }
}
