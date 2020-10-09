using System.IO;

namespace DoocumentManiipulateDemo
{
    public interface IDocumentMerger
    {
        /// <summary>
        /// edit Word document, convert to pdf 
        /// </summary>
        /// <returns>MemoryStream</returns>
        public MemoryStream TestDocument();
        /// <summary>
        /// Merge two documents into one
        /// </summary>
        /// <returns>byte[] of Merged Document</returns>
        byte[] PerformDocumentMerge();
    }
}