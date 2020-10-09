using Spire.Pdf;
using System.Collections.Generic;
using System.IO;

namespace DoocumentManiipulateDemo
{
    public static class Utils
    {
        /// <summary>
        /// Merge all Generated pdf agreement forms for an applicant
        /// </summary>
        /// <param name="pdfFilePaths"></param>
        /// <returns></returns>
        public static MemoryStream MergeMultiplePdfFiles(IReadOnlyList<MemoryStream> pdfFilePaths)
        {
            var count = pdfFilePaths.Count;

            if (count <= 0) return new MemoryStream();

            var docs = new PdfDocument[count];

            // open all documents
            for (var i = 0; i < count; i++)
                docs[i] = new PdfDocument(pdfFilePaths[i]);

            //merge all documents
            for (var i = 1; i < count; i++)
                docs[0].AppendPage(docs[i]);

            //save merged document to stream
            var ms = new MemoryStream();
            docs[0].SaveToStream(ms, FileFormat.PDF);

            //close all open documents
            foreach (var doc in docs) doc.Close();

            return ms;
        }
    }
}
