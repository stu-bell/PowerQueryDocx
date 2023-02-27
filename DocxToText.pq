// Custom Power Query function to extract data from a word docx file
// See comments in function DocxXmlExtractText
let
DocxToText = (Docx as binary) as text =>
    let Extract = ZipExtractAll(Docx),
        // Most of the text is included in the word/document.xml - but check out some of the other options here
        Xml = Xml.Document(Extract{[FileName="word/document.xml"]}[Content]),
        Text = DocxXmlExtractText(Xml)
    in Text,

DocxXmlExtractText = (DocxXml as table) as text => 
    // Recursive function. Parameter is a table returned by Xml.Document // https://github.com/stu0292/PowerQueryDocx
    Text.Combine(
        Table.AddColumn(DocxXml, "Text", each
            if Value.Type([Value]) = type text then [Value]
                // TODO handle different types of node by adding else if statements here
                // eg [Name] = "p" indicates a paragraph node, therefore insert new line before extracting the rest of the text 
                // For full reference see: https://learn.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document
            else if [Name] = "p" then "#(cr)#(lf)" & @DocxXmlExtractText([Value])
                else if [Name] = "tab" then "#(tab)" & @DocxXmlExtractText([Value])
                else @DocxXmlExtractText([Value])
        )[Text]
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
        
in DocxToText
