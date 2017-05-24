using System;
using System.Data;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;


namespace GenerateWordDoc
{
    /// <summary>
    /// Represents the complete sales report document to be generated.
    /// </summary>
    public class SalesReportBuilder
    {
        const string drawingTemplate = @"~/resources/drawingTemplate.xml";
        const string headerImageFile = @"~/resources/headerimage.gif";
        const string stylesXmlFile = @"~/resources/styles.xml";
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string wordPrefix = "w";

        private string _documentName;
        private string _salesPersonID;
        private string _imagePartRID;

        public SalesReportBuilder(string salesPersonID)
        {
            _salesPersonID = salesPersonID;
            //Generate unique filename
            _documentName = HttpContext.Current.Server.MapPath(
            @"~/reports/AdventureWorks" + DateTime.Now.ToFileTime() + ".docx");
        }

        #region Create Package and Parts
        /// <summary>
        /// 1. Create a new package as a Word document.
        /// 2. Add a style.xml part.
        /// 3. Add an embedded image part.
        /// 4. Create the document.xml part content. 
        /// </summary>
        /// <returns>File path location or error message</returns>
        public string CreateDocument()
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(_documentName, WordprocessingDocumentType.Document))
                {

                    // Set the content of the document so that Word can open it.
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

                    // Create a style part and add it to the document.
                    XmlDocument stylesXml = new XmlDocument();
                    stylesXml.Load(HttpContext.Current.Server.MapPath(stylesXmlFile));

                    StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    //  Copy the style.xml content into the new part....
                    using (Stream outputStream = stylePart.GetStream())
                    {
                        using (StreamWriter ts = new StreamWriter(outputStream))
                        {
                            ts.Write(stylesXml.InnerXml);
                        }
                    }

                    // Create an image part and add it to the document.
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Gif);
                    string imageFileName = System.Web.HttpContext.Current.Server.MapPath(headerImageFile);
                    using (FileStream stream = new FileStream(imageFileName, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }

                    // Get the reference ID for the image added to the package.
                    // You will use the image part reference ID to insert the
                    // image to the document.xml file.
                    _imagePartRID = mainPart.GetIdOfPart(imagePart);

                    // Create document.xml content.
                    SetMainDocumentContent(mainPart);
                }

                return (_documentName);
            }
            catch (Exception ex)
            {
                return (ex.Message);

            }
        }

        /// <summary>
        /// Set content of MainDocumentPart. 
        /// </summary>
        /// <param name="part">MainDocumentPart</param>
        public void SetMainDocumentContent(MainDocumentPart part)
        {
            using (Stream stream = part.GetStream())
            {
                CreateWordProcessingML(stream);
            }
        }

        /// <summary>
        /// Generate WordprocessingML for Sales Report.
        /// The resulting XML will be saved as document.xml.
        /// </summary>
        /// <param name="stream">MainDocumentPart stream</param>
        public void CreateWordProcessingML(Stream stream)
        {

            // Get sales person data from AdventureWorks database
            // You will write this data to the document.xml file.
            AdventureWorksSalesData salesData = new AdventureWorksSalesData();
            StringDictionary SalesPerson = salesData.GetSalesPersonData(_salesPersonID);

            // Create an XmlWriter using UTF8 encoding.
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.Indent = true;

            // This file represents the WordprocessingML of the Sales Report.
            XmlWriter writer = XmlWriter.Create(stream, settings);
            try
            {
                writer.WriteStartDocument(true);
                writer.WriteComment("This file represents the WordProcessingML of our Sales Report");
                writer.WriteStartElement(wordPrefix, "document", wordNamespace);
                writer.WriteStartElement(wordPrefix, "body", wordNamespace);

                WriteHeaderImage(writer);
                WriteDocumentTitle(writer, SalesPerson["FullName"]);
                WriteDocumentContactInfo(writer, SalesPerson["FullName"], SalesPerson["Phone"], SalesPerson["Email"]);
                WriteSalesSummaryInfo(writer, SalesPerson["SalesYTD"], SalesPerson["SalesQuota"]);
                WriteTerritoriesTable(writer, SalesPerson["TerritoryName"]);

                writer.WriteEndElement(); //body
                writer.WriteEndElement(); //document
            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                //Write the XML to file and close the writer.
                writer.Flush();
                writer.Close();
            }
            return;
        }
        #endregion


        #region Formatting Methods
        /// <summary>
        /// Write the title paragraph properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteTitleParagraphProperties(XmlWriter writer)
        {
            // Create the paragraph properties element.
            // </w:pPr>
            writer.WriteStartElement(wordPrefix, "pPr",
               wordNamespace);

            // Create the bottom border.
            //   <w:pBdr>
            //     <w:bottom w:val=”single” w:sz=”4” 
            //               w:space=”1” w:color=”auto” />
            //   </w:pBdr>
            writer.WriteStartElement(wordPrefix, "pBdr", wordNamespace);
            writer.WriteStartElement(wordPrefix, "bottom", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "single");
            writer.WriteAttributeString(wordPrefix, "sz", wordNamespace, "4");
            writer.WriteAttributeString(wordPrefix, "space", wordNamespace, "1");
            writer.WriteAttributeString(wordPrefix, "color", wordNamespace, "blue");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Define the spacing for the paragraph.
            //   <w:spacing w:line=”240” w:lineRule=”auto” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "line", wordNamespace, "240");
            writer.WriteAttributeString(wordPrefix, "lineRule", wordNamespace, "auto");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:pPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the title run properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteTitleRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);

            // Set up the spacing.
            //   <w:spacing w:val=”5” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "5");
            writer.WriteEndElement();

            // Define the size.
            //   <w:sz w:val=”52” />
            writer.WriteStartElement(wordPrefix, "sz", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "52");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:rPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the subtitle paragraph properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteSubtitleParagraphProperties(XmlWriter writer)
        {
            // Create the paragraph properties element.
            // <w:pPr>
            writer.WriteStartElement(wordPrefix, "pPr", wordNamespace);

            // Define the spacing for the paragraph.
            //   <w:spacing w:before=”200” w:after=”0” />
            writer.WriteStartElement(wordPrefix, "spacing", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "before", wordNamespace, "200");
            writer.WriteAttributeString(wordPrefix, "after", wordNamespace, "0");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:pPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the subtitle run properties to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteSubtitleRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);

            // setup as bold
            //   <w:b />
            writer.WriteElementString(wordPrefix, "b", wordNamespace, null);

            // Define the size.
            //   <sz w:val=”26” />
            writer.WriteStartElement(wordPrefix, "sz", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "26");
            writer.WriteEndElement();

            // Close the properties element.
            // </w:rPr>
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the bold run property to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the properties to.</param>
        private void WriteBoldRunProperties(XmlWriter writer)
        {
            // Create the run properties element.
            // <w:rPr>
            //   <w:b />
            // </w:rPr>
            writer.WriteStartElement(wordPrefix, "rPr", wordNamespace);
            writer.WriteElementString(wordPrefix, "b", wordNamespace, null);
            writer.WriteEndElement();
        }
        #endregion


        #region Styles Methods
        /// <summary>
        /// Write the style property to the WordprocessingML paragraph element.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the style to.</param>
        public void ApplyParagraphStyle(XmlWriter writer, string styleId)
        {
            // Apply the style in the paragraph properties.
            // <w:pPr>
            //   <w:pStyle w:val=”MyTitle” />
            // </w:pPr>
            writer.WriteStartElement(wordPrefix, "pPr", wordNamespace);
            writer.WriteStartElement(wordPrefix, "pStyle", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, styleId);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// Write the style property to the WordprocessingML table element.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the style to.</param>
        public void ApplyTableStyle(XmlWriter writer, string styleId)
        {
            // Apply the style in the table properties.
            // <w:tblPr>
            //   <w:tblStyle w:val="MyTableStyle" />
            //   <w:tblW w:w="0" w:type="auto" /> 
            //   <w:tblLook w:val="04A0" /> 
            // </w:tblPr>
            writer.WriteStartElement(wordPrefix, "tblPr", wordNamespace);
            writer.WriteStartElement(wordPrefix, "tblStyle", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, styleId);
            writer.WriteEndElement();
            writer.WriteStartElement(wordPrefix, "tblW", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "w", wordNamespace, "0");
            writer.WriteAttributeString(wordPrefix, "type", wordNamespace, "auto");
            writer.WriteEndElement();
            writer.WriteStartElement(wordPrefix, "tblLook", wordNamespace);
            writer.WriteAttributeString(wordPrefix, "val", wordNamespace, "04A0");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        #endregion


        #region document.xml writing methods
        /// <summary>
        /// Writes an image within a paragraph
        /// into the WordprocessingML.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the image to.</param>
        private void WriteHeaderImage(XmlWriter writer)
        {
            // Load the drawing template into an XML document.
            XmlDocument drawingXml = new XmlDocument();
            string drawingXmlFile = System.Web.HttpContext.Current.Server.MapPath(drawingTemplate);
            drawingXml.Load(drawingXmlFile);

            // Load the drawing template into an XML document and pass the reference ID parameter.
            drawingXml.LoadXml(string.Format(drawingXml.InnerXml, _imagePartRID));

            // Write the wrapping paragraph and the drawing fragment.
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            drawingXml.DocumentElement.WriteContentTo(writer);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a document title within a paragraph
        /// into the WordprocessingML.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the image to.</param>
        /// <param name="title">Document title</param>
        private void WriteDocumentTitle(XmlWriter writer, string title)
        {
            // Create the title.
            // <w:p>
            //   <w:r>
            //     <w:t>Sales Report - Employee Name</w:t>
            //   </w:r>
            // </w:p>    

            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            WriteTitleParagraphProperties(writer);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            WriteTitleRunProperties(writer);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Sales Report - " + title);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes a document's contact information within a paragraph
        /// into the WordprocessingML.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the image to.</param>
        /// <param name="fullname">Employee's fullname</param>
        /// <param name="phone">Employee's phone number</param>
        /// <param name="email">Employee's e-mail address</param>
        private void WriteDocumentContactInfo(XmlWriter writer, string fullname, string phone, string email)
        {

            // Create the contact information section.
            // <w:p>
            //   <w:r>
            //     <w:t>Contact</w:t>
            //   </w:r>
            // </w:p>

            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            WriteSubtitleParagraphProperties(writer);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            WriteSubtitleRunProperties(writer);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Contact");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Fill in the contact information section.
            // <w:p>
            //   <w:r>
            //     <w:t>Employee's fullname</w:t>
            //     <w:br />
            //     <w:t>sEmployee's e-mail</w:t>
            //     <w:br />
            //     <w:t>Employee's phone</w:t>
            //   </w:r>
            // </w:p>
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, fullname);
            writer.WriteElementString(wordPrefix, "br", wordNamespace, null);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, email);
            writer.WriteElementString(wordPrefix, "br", wordNamespace, null);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, phone);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes the sales summary information within a paragraph
        /// into the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the page break to.</param>
        /// <param name="totalsales">Employee's sales YTD</param>
        /// <param name="salesquota">Employee's sales quota</param>
        private void WriteSalesSummaryInfo(XmlWriter writer, string totalsales, string salesquota)
        {
            string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            string wordPrefix = "w";

            // Create the sales summary section.
            // <w:p>
            //   <w:r>
            //     <w:t>Sales Summary</w:t>
            //   </w:r>
            // </w:p>
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            WriteSubtitleParagraphProperties(writer);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            WriteSubtitleRunProperties(writer);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Sales Summary");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Fill in the contact information section.
            // <w:p>
            //   <w:r>
            //     <w:t>Total Sales:</w:t>
            //   </w:r>
            //   <w:r>
            //     <w:t>Employee's sales YTD</w:t>
            //   </w:r>
            //   <w:br />
            //   <w:r>
            //     <w:t>Employee's sales quota</w:t>
            //   </w:r>
            //   <w:r>
            //     <w:t>$1000.00</w:t>
            //   </w:r>
            // </w:p>
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Total Sales:");
            writer.WriteElementString(wordPrefix, "tab", wordNamespace, null);
            writer.WriteEndElement();
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, totalsales);
            writer.WriteEndElement();
            writer.WriteElementString(wordPrefix, "br", wordNamespace, null);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Sales Quota:");
            writer.WriteElementString(wordPrefix, "tab", wordNamespace, null);
            writer.WriteEndElement();
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, salesquota);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        /// <summary>
        /// Writes the territory sales totals as a table to the WordprocessingML document.
        /// </summary>
        /// <param name="writer">The XmlWriter to write the page break to.</param>
        private void WriteTerritoriesTable(XmlWriter writer, string territory)
        {
            string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            string wordPrefix = "w";

            // Create the territory section header.
            // <w:p>
            //   <w:r>
            //     <w:t>Sales by Territory</w:t>
            //   </w:r>
            // </w:p>
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            WriteSubtitleParagraphProperties(writer);
            ApplyParagraphStyle(writer, "Heading 3");
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            WriteSubtitleRunProperties(writer);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Sales by Territory - " + territory);
            writer.WriteEndElement();
            writer.WriteEndElement();

            // Open the table element.
            writer.WriteStartElement(wordPrefix, "tbl",
               wordNamespace);
            ApplyTableStyle(writer, "LightList-Accent2");

            // Write table headings.
            writer.WriteStartElement(wordPrefix, "tr", wordNamespace);
            writer.WriteStartElement(wordPrefix, "tc", wordNamespace);
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "Employee");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement(wordPrefix, "tc", wordNamespace);
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "2003");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement(wordPrefix, "tc", wordNamespace);
            writer.WriteStartElement(wordPrefix, "p", wordNamespace);
            writer.WriteStartElement(wordPrefix, "r", wordNamespace);
            writer.WriteElementString(wordPrefix, "t", wordNamespace, "2004");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            // write the close row
            writer.WriteEndElement();

            // write a row for each territory
            AdventureWorksSalesData salesData = new AdventureWorksSalesData();
            DataTable dt = salesData.GetSalesByTerritory(territory);

            foreach (DataRow myRow in dt.Rows)
            {
                writer.WriteStartElement(wordPrefix, "tr", wordNamespace);

                foreach (DataColumn myCol in dt.Columns)
                {
                    writer.WriteStartElement(wordPrefix, "tc", wordNamespace);
                    writer.WriteStartElement(wordPrefix, "p", wordNamespace);
                    writer.WriteStartElement(wordPrefix, "r", wordNamespace);
                    writer.WriteElementString(wordPrefix, "t", wordNamespace, myRow[myCol].ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                }
                // Write the close row.
                writer.WriteEndElement();
            }
            // end the table element
            writer.WriteEndElement();
        }

        #endregion

    }


}