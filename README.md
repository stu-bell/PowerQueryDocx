# PowerQueryDocx
Open Microsoft Word Docx files in Power Query

Copy and paste the code from the file DocxToText.pq into a new blank query in the power query editor. 
Then in another blank query, use the function something like this: 

```
let 
  FilePath = "C:/path/to/file.docx",
  Source = DocxToText(File.Contents(FilePath))
in Source
```

Microsoft Docs on custom Power Query functions: https://learn.microsoft.com/en-us/power-query/custom-function

Thanks to [Mark White](https://sql10.blogspot.com/2016/06/reading-zip-files-in-powerquery-m.html) for the zip extraction function
