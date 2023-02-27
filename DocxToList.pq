let 
// Custom Power Query function to extract data from a word docx file
// DocxToList(File.Contents("path/to/file.docx"))
// See comments in function DocxXmlExtractElements
DocxToList = (Docx as binary) as list =>
    let Extract = ZipExtractAll(Docx),
        // Most of the text is included in the word/document.xml - but check out some of the other options here
        Xml = Xml.Document(Extract{[FileName="word/document.xml"]}[Content]),
        Elements = DocxXmlExtractElements(Xml)
    in Elements,

DocxXmlExtractElements = (DocxXml as table) as list => 
    // Recursive function. Parameter is a table returned by Xml.Document // https://github.com/stu0292/PowerQueryDocx
    List.Combine(
        Table.AddColumn(DocxXml, "Elements", each
            if Value.Type([Value]) = type text then {[Value]} // text values must be inside a list
                // TODO handle different types of node by adding else if statements here
                // eg [Name] = "p" indicates a paragraph node, therefore insert new line before extracting the rest of the text 
                // For full reference see: https://learn.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document
                else if [Name] = "p" then {"#(cr)#(lf)"} & @DocxXmlExtractElements([Value])
                else if [Name] = "tab" then {"#(tab)"} & @DocxXmlExtractElements([Value])
                else @DocxXmlExtractElements([Value])
        )[Elements]
    ),

ZipExtractAll = (ZIPFile as binary) as table =>
    // source: http://sql10.blogspot.com/2016/06/reading-zip-files-in-powerquery-m.html
    let
        Header = BinaryFormat.Record([
            MiscHeader = BinaryFormat.Binary(14),
            BinarySize = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            FileSize   = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            FileNameLen= BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
            ExtrasLen  = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian)    
        ]),
        HeaderChoice = BinaryFormat.Choice(
            BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
            each if _ <> 67324752             // not the IsValid number? then return a dummy formatter
                then BinaryFormat.Record([IsValid = false, Filename=null, Content=null])
                else BinaryFormat.Choice(
                        BinaryFormat.Binary(26),      // Header payload - 14+4+4+2+2
                        each BinaryFormat.Record([
                            IsValid  = true,
                            Filename = BinaryFormat.Text(Header(_)[FileNameLen]),
                            Extras   = BinaryFormat.Text(Header(_)[ExtrasLen]),
                            Content  = BinaryFormat.Transform(
                                BinaryFormat.Binary(Header(_)[BinarySize]),
                                (x) => try Binary.Buffer(Binary.Decompress(x, Compression.Deflate)) otherwise null
                            )
                            ]),
                            type binary                   // enable streaming
                    )
        ),
        ZipFormat = BinaryFormat.List(HeaderChoice, each _[IsValid] = true),
        Entries = List.Transform(
            List.RemoveLastN( ZipFormat(ZIPFile), 1),
            (e) => [FileName = e[Filename], Content = e[Content] ]
        )
    in
        Table.FromRecords(Entries)
        
in DocxToList
