<div align="center">

## Memory Mapped File Wrapper Class


</div>

### Description

A Memory Mapped File resides in virtual memory from the moment it is opened until it is closed.

This makes access to it very fast, in fact read and write merely consist of data moves in memory.

All parts of the file are accessible by means of a zero based offset and the length of

the data chunks you can transfer is only limited by the file size itself. The system will

swap out and in pages of virtual memory as necessary.

Only when you close the file the data are "lazily" written to disc, and the Physical

Write Operations are limited to the altered memory pages.

When you open an MMF you are required to supply an estimate of how large the file is gonna

be. Any reasonable estimate will be fine, however, you cannot exceed this estimate, ie you

can't extend a file while it is open, but you can truncate the file size to the actual

size used when you close it. You can extend an existing file by giving the appropriate

estimate when you open it and truncating it to the larger size when closing.

MMFs are byte oriented and that requires a bit of understanding the data types which VB uses

and how they are represented physically. The class contains tools however to handle the

different aspects.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-03-17 14:04:26
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Advanced
**User Rating**    |4.9 (78 globes from 16 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Memory\_Map1561063172003\.zip](https://github.com/Planet-Source-Code/ulli-memory-mapped-file-wrapper-class__1-44076/archive/master.zip)

### API Declarations

a few...





