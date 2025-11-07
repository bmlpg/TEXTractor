using Toxy;
using Toxy.Parsers;

namespace TEXTractor
{
    public class TEXTractor : ITEXTractor
    {
        public void GetVCard(string FileName, byte[] FileContent, out BusinessCards ssBusinessCards)
        {
            ParserContext context = new ParserContext(FileName, FileContent);
            VCardDocumentParser parser = ParserFactory.CreateVCard(context);
            ToxyBusinessCards toxyBusinessCards = parser.Parse();

            BusinessCards businessCards = new BusinessCards();
            businessCards.Cards = new List<BusinessCard>();

            foreach (ToxyBusinessCard item in toxyBusinessCards.Cards)
            {
                BusinessCard businessCard = new BusinessCard();

                if (item.Name != null)
                {
                    businessCard.Name.FirstName = item.Name.FirstName;
                    businessCard.Name.MiddleName = item.Name.MiddleName;
                    businessCard.Name.LastName = item.Name.LastName;
                    businessCard.Name.FullName = item.Name.FullName;
                }

                if (item.NickName != null)
                {
                    businessCard.NickName.FirstName = item.NickName.FirstName;
                    businessCard.NickName.MiddleName = item.NickName.MiddleName;
                    businessCard.NickName.LastName = item.NickName.LastName;
                    businessCard.NickName.FullName = item.NickName.FullName;
                }

                businessCard.Organization = item.Organization;
                businessCard.Title = item.Title;
                businessCard.Class = item.Class;
                businessCard.TimeZone = item.TimeZone;
                businessCard.UID = item.UID;
                businessCard.GEO = item.GEO;
                businessCard.Label = item.Label;

                businessCard.Categories = new List<Text>();
                if (item.Categories != null)
                {
                    foreach (string item1 in item.Categories)
                    {
                        Text category = new Text();
                        category.Value = item1;
                        businessCard.Categories.Add(category);
                    }
                }

                businessCard.Birthday = item.Birthday;

                businessCard.Photos = new List<Text>();
                if (item.Photos != null)
                {
                    foreach (string item1 in item.Photos)
                    {
                        Text photo = new Text();
                        photo.Value = item1;
                        businessCard.Photos.Add(photo);
                    }
                }

                businessCard.Sources = new List<Text>();
                if (item.Sources != null)
                {
                    foreach (string item1 in item.Sources)
                    {
                        Text source = new Text();
                        source.Value = item1;
                        businessCard.Sources.Add(source);
                    }
                }

                businessCard.ProductID = item.ProductID;

                businessCard.Addresses = new List<Address>();
                if (item.Addresses != null)
                {
                    foreach (ToxyAddress item1 in item.Addresses)
                    {
                        Address address = new Address();
                        address.Street = item1.Street;
                        address.City = item1.City;
                        address.Region = item1.Region;
                        address.PostalCode = item1.PostalCode;
                        address.Country = item1.Country;
                        businessCard.Addresses.Add(address);
                    }
                }

                businessCard.Contacts = new List<Contact>();
                if (item.Contacts != null)
                {
                    foreach (ToxyContact item1 in item.Contacts)
                    {
                        Contact contact = new Contact();
                        contact.Name = item1.Name;
                        contact.Value = item1.Value;
                        businessCard.Contacts.Add(contact);
                    }
                }

                if (item.Gender != GenderType.Unknown)
                    businessCard.Gender = item.Gender.ToString();

                businessCards.Cards.Add(businessCard);
            }

            ssBusinessCards = businessCards;

        }

        public void GetDom(string FileName, byte[] FileContent, out Dom ssDom)
        {
            ParserContext context = new ParserContext(FileName, FileContent);
            IDomParser parser = ParserFactory.CreateDom(context);
            ToxyDom toxyDom = parser.Parse();

            Dom dom = new Dom();
            dom.Nodes = new List<Node>();

            CopyNodes(toxyDom.Root, dom, "", "1");

            ssDom = dom;
        }

        private void CopyNodes(ToxyNode toxyNode, Dom dom, string parentIdentifier, string ownIdentifier)
        {
            Node node = new Node();
            node.Name = toxyNode.Name;
            node.Text = toxyNode.Text;
            node.ParentIdentifier = parentIdentifier;
            node.Identifier = ownIdentifier;

            node.Attributes = new List<Attribute>();
            foreach (ToxyAttribute item in toxyNode.Attributes)
            {
                Attribute attribute = new Attribute();
                attribute.Name = item.Name;
                attribute.Value = item.Value;

                node.Attributes.Add(attribute);
            }

            dom.Nodes.Add(node);

            int index = 1;
            foreach (ToxyNode item in toxyNode.ChildrenNodes)
            {
                CopyNodes(item, dom, ownIdentifier, ownIdentifier + "." + index);
                index++;
            }
        }

        public void GetSlideshow(string FileName, byte[] FileContent, out Slideshow ssSlideshow)
        {
            ParserContext context = new ParserContext(FileName, FileContent);
            ISlideshowParser parser = ParserFactory.CreateSlideshow(context);
            ToxySlideshow toxySlideshow = parser.Parse();

            Slideshow slideshow = new Slideshow();

            slideshow.Slides = new List<Slide>();
            foreach (ToxySlide item in toxySlideshow.Slides)
            {
                Slide slide = new Slide();

                slide.Texts = new List<Text>();
                foreach (string item1 in item.Texts)
                {
                    Text slideText = new Text();
                    slideText.Value = item1;

                    slide.Texts.Add(slideText);
                }

                slide.Note = item.Note;

                slideshow.Slides.Add(slide);
            }

            ssSlideshow = slideshow;
        }

        public void GetSpreadsheet(string FileName, byte[] FileContent, SpreadsheetCSVOptions CSVOptions, SpreadsheetExcelOptions ExcelOptions, out Spreadsheet ssSpreadsheet)
        {
            ParserContext context = new ParserContext(FileName, FileContent);

            context.Properties.Add("ExtractHeader", (!CSVOptions.HasNoHeader).ToString());
            if (CSVOptions.Delimeter != "")
                context.Properties.Add("Delimiter", CSVOptions.Delimeter);

            context.Properties.Add("HasHeader", ExcelOptions.HasHeader.ToString());
            context.Properties.Add("ExtractSheetHeader", ExcelOptions.ExtractSheetHeader.ToString());
            context.Properties.Add("ExtractSheetFooter", ExcelOptions.ExtractSheetFooter.ToString());
            context.Properties.Add("ShowCalculatedResult", ExcelOptions.ShowCalculatedResult.ToString());
            context.Properties.Add("FillBlankCells", ExcelOptions.FillBlankCells.ToString());
            context.Properties.Add("IncludeComments", (!ExcelOptions.DoNotIncludeComments).ToString());

            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet toxySpreadsheet = parser.Parse();

            Spreadsheet spreadsheet = new Spreadsheet();

            spreadsheet.Name = toxySpreadsheet.Name;
            spreadsheet.Length = toxySpreadsheet.Length;
            spreadsheet.Tables = new List<Table>();

            foreach (ToxyTable item in toxySpreadsheet.Tables)
            {
                Table table = new Table();

                table.Name = item.Name;
                table.HasHeader = item.HasHeader;
                table.LastColumnIndex = item.LastColumnIndex;
                table.LastRowIndex = item.LastRowIndex;
                table.Length = item.Length;
                table.PageHeader = item.PageHeader;
                table.PageFooter = item.PageFooter;

                table.HeaderRows = ConvertToxyTableRows(item.HeaderRows);

                table.Rows = ConvertToxyTableRows(item.Rows);

                table.MergeCells = new List<MergeCellRange>();
                foreach (Toxy.MergeCellRange item1 in item.MergeCells)
                {
                    MergeCellRange mergeCellRange = new MergeCellRange();

                    mergeCellRange.FirstRow = item1.FirstRow;
                    mergeCellRange.LastRow = item1.LastRow;
                    mergeCellRange.FirstColumn = item1.FirstColumn;
                    mergeCellRange.LastColumn = item1.LastColumn;

                    table.MergeCells.Add(mergeCellRange);
                }

                spreadsheet.Tables.Add(table);
            }

            ssSpreadsheet = spreadsheet;
        }

        private List<Row> ConvertToxyTableRows(List<ToxyRow> toxyRows)
        {
            List<Row> rows = new List<Row>();

            foreach (ToxyRow item in toxyRows)
            {
                Row row = new Row();

                row.RowIndex = item.RowIndex;
                row.Length = item.Length;
                row.LastCellIndex = item.LastCellIndex;

                row.Cells = new List<Cell>();
                foreach (ToxyCell item1 in item.Cells)
                {
                    Cell cell = new Cell();

                    cell.Value = item1.Value;
                    cell.CellIndex = item1.CellIndex;
                    cell.Comment = item1.Comment;
                    cell.Formula = item1.Formula;

                    row.Cells.Add(cell);
                }

                rows.Add(row);
            }

            return rows;
        }

        public void GetEmail(string FileName, byte[] FileContent, out Email ssEmail)
        {
            ParserContext context = new ParserContext(FileName, FileContent);
            IEmailParser parser = ParserFactory.CreateEmail(context);
            ToxyEmail toxyEmail = parser.Parse();

            Email email = new Email();
            email.From = toxyEmail.From;

            email.To = new List<Text>();
            foreach (string item in toxyEmail.To)
            {
                Text emailAddress = new Text();
                emailAddress.Value = item;

                email.To.Add(emailAddress);
            }

            email.Cc = new List<Text>();
            foreach (string item in toxyEmail.Cc)
            {
                Text emailAddress = new Text();
                emailAddress.Value = item;

                email.Cc.Add(emailAddress);
            }

            email.Bcc = new List<Text>();
            foreach (string item in toxyEmail.Bcc)
            {
                Text emailAddress = new Text();
                emailAddress.Value = item;

                email.Bcc.Add(emailAddress);
            }

            email.Attachments = new List<Text>();
            foreach (string item in toxyEmail.Attachments)
            {
                Text attachment = new Text();
                attachment.Value = item;

                email.Attachments.Add(attachment);
            }

            email.HtmlBody = toxyEmail.HtmlBody;

            email.Subject = toxyEmail.Subject;

            email.TextBody = toxyEmail.TextBody;

            if (toxyEmail.ArrivalTime.HasValue)
            {
                email.ArrivalTime = (DateTime)toxyEmail.ArrivalTime;
            }

            ssEmail = email;

        }

        public void GetText(string FileName, byte[] FileContent, TextExcelOptions ExcelOptions, TextWordOptions WordOptions, out string ssText)
        {
            ParserContext context = new ParserContext(FileName, FileContent);

            context.Properties.Add("IncludeHeaderFooter", ExcelOptions.IncludeHeaderFooter.ToString());
            context.Properties.Add("ShowCalculatedResult", ExcelOptions.ShowCalculatedResult.ToString());
            context.Properties.Add("IncludeSheetNames", (!ExcelOptions.DoNotIncludeSheetNames).ToString());
            context.Properties.Add("IncludeComments", (!ExcelOptions.DoNotIncludeComments).ToString());

            context.Properties.Add("ExtractHeader", WordOptions.ExtractHeader.ToString());
            context.Properties.Add("ExtractFooter", WordOptions.ExtractFooter.ToString());

            ITextParser parser = ParserFactory.CreateText(context);
            string text = parser.Parse();

            ssText = text;
        }

        public void GetMetadata(string FileName, byte[] FileContent, out Metadata ssMetadata)
        {
            ParserContext context = new ParserContext(FileName, FileContent);
            IMetadataParser parser = ParserFactory.CreateMetadata(context);
            ToxyMetadata toxyMetadata = parser.Parse();

            Metadata metadata = new Metadata();
            metadata.Properties = new List<Property>();

            IEnumerator<ToxyProperty> enumerator = toxyMetadata.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Property property = new Property();
                property.Name = enumerator.Current.Name;
                property.Value = enumerator.Current.Value.ToString();

                metadata.Properties.Add(property);
            }

            ssMetadata = metadata;
        }

        public void GetDocument(string FileName, byte[] FileContent, DocumentWordOptions WordOptions, out Document ssDocument)
        {
            ParserContext context = new ParserContext(FileName, FileContent);

            context.Properties.Add("ExtractHeader", WordOptions.ExtractHeader.ToString());
            context.Properties.Add("ExtractFooter", WordOptions.ExtractFooter.ToString());

            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument toxyDocument = parser.Parse();

            Document document = new Document();
            document.Header = toxyDocument.Header;
            document.Footer = toxyDocument.Footer;
            document.TotalPageNumber = toxyDocument.TotalPageNumber;
            document.Paragraphs = new List<Paragraph>();

            foreach (ToxyParagraph toxyParagraph in toxyDocument.Paragraphs)
            {
                Paragraph paragraph = new Paragraph();
                paragraph.Text = toxyParagraph.Text;
                paragraph.StyleID = toxyParagraph.StyleID;

                if (toxyParagraph.EmbededStyle != null)
                {
                    ParagraphStyle embededStyle = new ParagraphStyle();
                    embededStyle.IsBold = toxyParagraph.EmbededStyle.IsBold;
                    embededStyle.IsItalic = toxyParagraph.EmbededStyle.IsItalic;
                    embededStyle.IsUnderlined = toxyParagraph.EmbededStyle.IsUnderlined;
                    embededStyle.FontFamily = toxyParagraph.EmbededStyle.FontFamily;
                    embededStyle.FontSize = toxyParagraph.EmbededStyle.FontSize;
                    paragraph.EmbededStyle = embededStyle;
                }

                document.Paragraphs.Add(paragraph);
            }

            ssDocument = document;
        }

    }
}
