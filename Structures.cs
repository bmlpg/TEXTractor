using OutSystems.ExternalLibraries.SDK;

namespace TEXTractor
{
    [OSStructure(Description = "")]
    public struct Address
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Street;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string City;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Region;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string PostalCode;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Country;
    }

    [OSStructure(Description = "")]
    public struct Attribute
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Value;
    }

    [OSStructure(Description = "")]
    public struct BusinessCard
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public Name Name;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public Name NickName;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Organization;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Title;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Class;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string TimeZone;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string UID;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string GEO;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Label;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Categories;
        [OSStructureField(DataType = OSDataType.DateTime, Description = "", IsMandatory = true)]
        public DateTime Birthday;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Photos;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Sources;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string ProductID;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Address> Addresses;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Contact> Contacts;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Gender;
    }

    [OSStructure(Description = "")]
    public struct BusinessCards
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<BusinessCard> Cards;
    }

    [OSStructure(Description = "")]
    public struct Cell
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Value;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int CellIndex;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Comment;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Formula;
    }

    [OSStructure(Description = "")]
    public struct Contact
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Value;
    }

    [OSStructure(Description = "")]
    public struct Document
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Header;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Footer;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Paragraph> Paragraphs;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int TotalPageNumber;
    }

    [OSStructure(Description = "")]
    public struct DocumentWordOptions
    {
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractHeader;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractFooter;
    }

    [OSStructure(Description = "")]
    public struct Dom
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Node> Nodes;
    }

    [OSStructure(Description = "")]
    public struct Email
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string From;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> To;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Cc;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Bcc;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Attachments;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string HtmlBody;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Subject;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string TextBody;
        [OSStructureField(DataType = OSDataType.DateTime, Description = "", IsMandatory = true)]
        public DateTime ArrivalTime;
    }

    [OSStructure(Description = "")]
    public struct MergeCellRange
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int FirstColumn;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int FirstRow;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int LastColumn;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int LastRow;
    }

    [OSStructure(Description = "")]
    public struct Metadata
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Property> Properties;
    }

    [OSStructure(Description = "")]
    public struct Name
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string FirstName;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string MiddleName;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string LastName;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string FullName;

    }

    [OSStructure(Description = "")]
    public struct Node
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string ParentIdentifier;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Identifier;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Text;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Attribute> Attributes;
    }

    [OSStructure(Description = "")]
    public struct Paragraph
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Text;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string StyleID;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public ParagraphStyle EmbededStyle;
    }

    [OSStructure(Description = "")]
    public struct ParagraphStyle
    {
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = true)]
        public bool IsBold;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = true)]
        public bool IsItalic;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = true)]
        public bool IsUnderlined;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int FontSize;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string FontFamily;
    }

    [OSStructure(Description = "")]
    public struct Property
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Value;
    }

    [OSStructure(Description = "")]
    public struct Row
    {
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int RowIndex;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int LastCellIndex;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int Length;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Cell> Cells;
    }

    [OSStructure(Description = "")]
    public struct Slide
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Text> Texts;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Note;
    }

    [OSStructure(Description = "")]
    public struct Slideshow
    {
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Slide> Slides;
    }

    [OSStructure(Description = "")]
    public struct Spreadsheet
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Table> Tables;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int Length;
    }

    [OSStructure(Description = "")]
    public struct SpreadsheetCSVOptions
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "Content delimeter. Default is \",\".", IsMandatory = false)]
        public string Delimeter;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool HasNoHeader;
    }

    [OSStructure(Description = "")]
    public struct SpreadsheetExcelOptions
    {
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool HasHeader;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractSheetHeader;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractSheetFooter;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ShowCalculatedResult;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool FillBlankCells;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool DoNotIncludeComments;
    }

    [OSStructure(Description = "")]
    public struct Table
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Name;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Row> HeaderRows;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<Row> Rows;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int LastColumnIndex;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int LastRowIndex;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string PageHeader;
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string PageFooter;
        [OSStructureField(DataType = OSDataType.InferredFromDotNetType, Description = "", IsMandatory = true)]
        public List<MergeCellRange> MergeCells;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = true)]
        public bool HasHeader;
        [OSStructureField(DataType = OSDataType.Integer, Description = "", IsMandatory = true)]
        public int Length;
    }

    [OSStructure(Description = "")]
    public struct Text
    {
        [OSStructureField(DataType = OSDataType.Text, Description = "", IsMandatory = true)]
        public string Value;
    }

    [OSStructure(Description = "")]
    public struct TextExcelOptions
    {
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool IncludeHeaderFooter;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ShowCalculatedResult;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool DoNotIncludeSheetNames;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool DoNotIncludeComments ;
    }

    [OSStructure(
        Description = ""
    )]
    public struct TextWordOptions
    {
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractHeader;
        [OSStructureField(DataType = OSDataType.Boolean, Description = "", IsMandatory = false)]
        public bool ExtractFooter;
    }

}
