# Doc/DocX to HTML Converter
 Converts Word documents into clean HTML

I have this problem: mo matter what my official job or title, people keep sending me Word documents that they want posted online to match the web site styling.

Yes, you can use Word to convert documents to HTML, but Microsoft's version of "HTML" frequently looks worse than if you just pasted in plain text. And yes, you can save Word documents as plain text, but then to use them on the web you have to add in the HTML tags.

Finally I got fed up and wrote a converter to produce minimally formatted HTML that I can copy into common web editors like CKEditor or TinyMCE. The operation is linear: you take a Word .doc or .docx file, drag it onto a Windows form, the program invokes Word, converts your .doc/.docx to the cleanest HTML that Word can manage, parses the HTML using [Html Agility Pack](https://github.com/zzzprojects/html-agility-pack), and finally spits out a simple HTML document in Notepad that you can copy-paste into whatever web system you're using.

Prerequisites:

1. When building this I had the Microsoft.Office.Interop.Word 12.0 (Word 2007) library referenced from the project. The easiest way to meet this requirement is to install some recent version of Office, but any version of the Microsoft.Office.Interop.Word library that natively handles the .docx format should work.
2. I had the very useful Html Agility Pack version 1.4.6 library referenced as well. I was using .NET 4.0 and the 4.0 version of the library. Html Agility Pack now [lives on Github](https://github.com/zzzprojects/html-agility-pack) so you can grab it easily and reference it from your project. 

Later I realized a better way to do this might be to invoke the LibreOffice converter on the command line, convert your document to HTML or text, filter it with Python's BeautifulSoup library or sed or Ruby's Nokogiri, and then insert the results straight into the database of your web system. But maybe not: in text, tags like <<span>table</span>> and <<span>ul</span>> would be lost, and LibreOffice's HTML is still pretty ugly.

License
---------------
GNU General Public License v3.0

Author
---------------
Geoff Stratton
