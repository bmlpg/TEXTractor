using OutSystems.ExternalLibraries.SDK;

namespace TEXTractor
{
    
    [OSInterface(Name = "TEXTractor", Description = "Extract text and/or metadata from 32 different file formats, including pdf, xlsx, csv, and docx.", IconResourceName = "TEXTractor.resources.textractor_logo.png")]
    public interface ITEXTractor
    {
        [OSAction(Description = "Get vcard content in a structured format.")]
        public void GetVCard(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: vcf.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Structured representation of the vcard content.")]
            out BusinessCards BusinessCards
        );

        [OSAction(Description = "Get dom content in a structured format.")]
        public void GetDom(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: htm, html, xml.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Structured representation of the dom content.")]
            out Dom Dom
        );

        [OSAction(Description = "Get slideshow content in a structured format.")]
        public void GetSlideshow(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: pptx.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Structured representation of the slideshow content.")]
            out Slideshow Slideshow
        );

        [OSAction(Description = "Get spreadsheet content in a structured format.")]
        public void GetSpreadsheet(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: csv, xls, xlsx.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "CSV extraction options.")]
            SpreadsheetCSVOptions CSVOptions,
            [OSParameter(Description = "Excel extraction options.")]
            SpreadsheetExcelOptions ExcelOptions,
            [OSParameter(Description = "Structured representation of the spreadsheet content.")]
            out Spreadsheet Spreadsheet
        );

        [OSAction(Description = "Get email content in a structured format.")]
        public void GetEmail(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: eml, msg.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Structured representation of the email content.")]
            out Email Email
        );
       
        [OSAction(Description = "Get file content in plain text.")]
        public void GetText(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: csv, docx, eml, htm, html, msg, pdf, pptx, rtf, txt, vcf, xls, xlsx, xml, zip.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Excel extraction options.")]
            TextExcelOptions ExcelOptions,
            [OSParameter(Description = "Word extraction options.")]
            TextWordOptions WordOptions,
            [OSParameter(Description = "Text representation of the file content.")]
            out string Text
        );

        [OSAction(Description = "Get file metadata in a structured format.")]
        public void GetMetadata(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: aif, ape, docx, flac, gif, jpeg, jpg, pubx, mp3, png, ppt, pptx, pub, shw, sldprt, tiff, vsd, vsdx, wma, xls, xlsx.")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Structured representation of the file metadata.")]
            out Metadata Metadata
        );

        [OSAction(Description = "Get document content in a structured format.")]
        public void GetDocument(
            [OSParameter(Description = "File name including extension. Supported extensions/file types: docx, pdf")]
            string FileName,
            [OSParameter(Description = "Binary content of the file.")]
            byte[] FileContent,
            [OSParameter(Description = "Word extraction options.")]
            DocumentWordOptions WordOptions,
            [OSParameter(Description = "Structured representation of the document content.")]
            out Document Document
        );

    }    
}
