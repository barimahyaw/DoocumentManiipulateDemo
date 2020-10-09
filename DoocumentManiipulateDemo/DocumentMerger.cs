using System;
using Spire.Doc;
using System.IO;
using Microsoft.AspNetCore.Hosting;

namespace DoocumentManiipulateDemo
{
    public class DocumentMerger : IDocumentMerger
    {
        private readonly IWebHostEnvironment _hostEnvironment;

        public DocumentMerger(IWebHostEnvironment hostEnvironment)
        {
            _hostEnvironment = hostEnvironment;
        }
        /// <summary>
        /// edit Word document, convert to pdf 
        /// </summary>
        /// <returns>MemoryStream</returns>
        public MemoryStream TestDocument()
        {
            const string guarantor = " ";

            // Get Path of Template of Word/Pdf/Excel/Image/Html file from your local drive
            var path = Path.Combine(_hostEnvironment.WebRootPath, "Templates/GuarantorReleasedCertificate.docx");


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
            return ms;
        }

        /// <summary>
        ///  edit first Word document, convert to pdf and save it to Stream
        /// </summary>
        /// <returns>MemoryStream</returns>
        private MemoryStream FirstDocument()
        {
            const string guarantor = " ";

            // Get Path of Template of Word/Pdf/Excel/Image/Html file from your local drive
            var path = Path.Combine(_hostEnvironment.WebRootPath, "Templates/GuarantorReleasedCertificate.docx");


            using var document = new Document(path);

            // Perform the replacement of placeholders:    
            var guarantorName = !string.IsNullOrWhiteSpace(guarantor)
                ? "Guarantor Institution Name"
                : "Philip Aseidu";
            const string borrowerName = "Mark Afari";

            document.Replace("%DatePrepared%", DateTime.Now.ToString("D"), false, true);
            document.Replace("%GuarantorName%", guarantorName, false, true);
            document.Replace("%GuarantorSsf%", "B178299150013", false, true);
            document.Replace("%GuarantorPostalAddress%", "Ghana Education Service", false, true);
            document.Replace("%BorrowerName%", borrowerName, false, true);
            document.Replace("%BorrowerSsf%", "C018512310052", false, true);
            document.Replace("%BorrowerPostalAddress%", "P. O. Box 5, Koforidua", false, true);
            document.Replace("%Institution%", "Koforidua Technical University", false, true);
            document.Replace("%GuarantorReleasedBy%", "Emmanuel Agyei", false, true);
            document.Replace("%Destination%", "Administrator", false, true);
            document.Replace("%OfficeBranch%", "Accra Main", false, true);
            document.Replace("%OfficeBranch%", "Accra Main", false, true);

            var ms = new MemoryStream();

            document.SaveToStream(ms, FileFormat.PDF);

            //specify the position of file in the stream to use for seek
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }
        /// <summary>
        ///  edit second Word document, convert to pdf and save it to Stream
        /// </summary>
        /// <returns>MemoryStream</returns>
        private MemoryStream SecondDocument()
        {
            const string guarantor = " ";

            // Get Path of Template of Word/Pdf/Excel/Image/Html file from your local drive
            var path = Path.Combine(_hostEnvironment.WebRootPath, "Templates/GuarantorReleasedCertificate.docx");


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

            var ms = new MemoryStream();

            document.SaveToStream(ms, FileFormat.PDF);

            //specify the position of file in the stream to use for seek
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }
        /// <summary>
        /// Merge two documents into one
        /// </summary>
        /// <returns>byte[] of Merged Document</returns>
        public byte[] PerformDocumentMerge()
        {
            var memoryStreams = new MemoryStream[2];

            memoryStreams[0] = FirstDocument();

            memoryStreams[1] = SecondDocument();

            var mergedDocument = Utils.MergeMultiplePdfFiles(memoryStreams).ToArray();

            return mergedDocument;
        }
    }
}
