# WPS Office Adapter

Greetings AutoIt community!

After doing a quick search and finding nothing posted for working with WPS Office using AutoIt, I decided to whip up the following small set of functions.

In my testing and research I have found that WPS Office seems to be just a free, re-branded version of MSOffice. Underneath the makeup, the COM methods etc are all the same, even the CLSID’s! That being the case you should be able to use the UDF’s by water (or others) meant for working with Microsoft Office’s products, such as the Word and Excel UDF’s included with AutoIt already, as well as the PowerPoint UDF by water, found here: <https://www.autoitscript.com/forum/files/file/374-powerpoint-udf/>

I have not exhaustively tested automating WPS using the MSWord COM methods, though I have used many MSOffice COM methods, etc., for automating WPS without a problem, you are encouraged to do testing yourself and make sure it works. I have minimal experience with MSOffice and related products, and thus may not be able to help much with bugs, but I will try to help as time permits.

The question will arise, why did I have to make these functions, if these two Office programs are seemingly identical?
1. To open a new instance of WPS Office (Writer, Spreadsheets, Presentation), the Class names are different. (KWPS.Application vs Word.Application; KWPP.Application vs PowerPoint.Application; KET.Application vs Excel.Application). And,
2. When connecting (or attaching) to an existing instance of WPS, you can easily get a MSOffice application Object instead of a WPS Application Object, and vice versa, if you so happen to have both installed and running. Thus I added a check for the Registry path to contain the string “\kingsoft\” to ensure you retrieve the intended object.

The functions I made are copies from those written by water in the respective MS Office UDFs, with modifications to work with WPS, and some personal preference. Once you retrieve the Object you need for the WPS application, go forward using those Objects in their respective MS Office UDF functions.

Question or comments, or suggestions are welcome and appreciated. And also, if you wish to take this function and improve it or expand it, feel free, I’m not intending to go further than I have, with the exception of perhaps bug fixes as time permits.

## Notes

- If anyone has a better method for connecting to and differentiating between an active instance of WPS, and an active instance of MS Office, I would appreciate any hints. The CLSID’s are identical, as well as the class name when it is running (Word.Application; Excel.Application; PowerPoint.Application)
- So far I haven’t come up with a better name, so I have called it simply “WPS Office Adapter” because it allows you to use MS Office UDF functions for WPS. Any suggestions for a name will also be appreciated.
- In order for the examples of WPS Presentation to work, you will need to modify the Include path to point to the PowerPoint UDF, or place a copy into the main WPS UDF folder.
- For the `_Excel_Close` function, you will need to call ForceClose with True, or implement water’s internal `__Excel_CloseOnQuit` function as done in _Excel_Open, to allow the usual functionality of `_Excel_Close` to work.

## Equivalent functions

The Functions that replace existing MS Office functions are as follows:

|         WPS Function         |   MS Office Equivalent   |
|:----------------------------:|:------------------------:|
| _WPS_Presentation_Attach     | _PPT_PresentationAttach  |
| _WPS_Presentation_Open       | _PPT_Open                |
| _WPS_Spreadsheets_BookAttach | _Excel_BookAttach        |
| _WPS_Spreadsheets_BookList   | _Excel_BookList          |
| _WPS_Spreadsheets_Open       | _Excel_Open              |
| _WPS_Writer_Create           | _Word_Create             |
