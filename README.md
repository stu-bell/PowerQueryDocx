# PowerQueryDocx
Open Microsoft Word Docx files in Power Query

Copy and paste the code from the file [DocxToListpq](https://raw.githubusercontent.com/stu0292/PowerQueryDocx/main/DocxToList.pq) into a new blank query in the power query editor. 
Then in another blank query, use the function something like this: 

```
let 
  FilePath = "C:/path/to/file.docx",
  Source = DocxToList(File.Contents(FilePath))
in Source
```

Microsoft Docs on custom Power Query functions: https://learn.microsoft.com/en-us/power-query/custom-function

This version is very basic and only picks out text and new lines. If there are specific elements in your word doc you want to extract, you'll need to figure out the element names and attributes from the [Open XML docs](https://learn.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document) and modify the extract text function.

Thanks to [Mark White](https://sql10.blogspot.com/2016/06/reading-zip-files-in-powerquery-m.html) for the zip extraction function
