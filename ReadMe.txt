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

It is up to you whether you want to use and react to all the error returns. A zero return
indicates success in all cases.


Thanks go to Randy Birch, Eduardo A. Morcillo, and Karl E. Peterson for inspiring this class.

Sorry about the StrPtr Vlad - my fault :-)