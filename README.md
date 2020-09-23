<div align="center">

## Commas Seperated Values Files \- Class Module


</div>

### Description

It's easy to generate a CSV file with any delimiter by using this class module.

Hope useful to anyone.
 
### More Info
 
How to use it?

See sample

First, add this module to class module entry in your project.

Declare a variable:

Dim CSVFile as New NewCSV

With CSVFile

.FileName=

.Delimiter={input one character}

.Header={default is False}

If Dir(.FileName)&lt;&gt;"" then

.Open

.MoveFirst

Do While Not .EOfF

Debug.Print .Fields(0)

.MoveNext

Loop

EndIf

End With


<span>             |<span>
---                |---
**Submitted On**   |2005-06-03 15:32:34
**By**             |[Nguyen Thanh Canh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nguyen-thanh-canh.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Commas\_Sep19600112272005\.zip](https://github.com/Planet-Source-Code/nguyen-thanh-canh-commas-seperated-values-files-class-module__1-63809/archive/master.zip)








