# What are these scripts?

I was asked by a friend of mine to help him create some tools for docx modification and pdf creation for bureaucratic work. Here are some simplified, barebone templates and experiments creating some docx and pdf conversors in Python. They can be easily modified for specific applications. In general, working with .docx files is a complete pain and working with .pdf files is mostly fine. The tools for .docx modification in Python require the user to either have installed Microsoft Word or LibreOffice Writer or use some proprietary Python package:

* `pip install python-docx`: The python-docx is a very unintuitive package, but it gets the job done. If you try to get exact copies of the style and text of another .docx file, chances are you are getting a "mostly" similar output file with some obscure differences. Luckily, it does not requiere neither MS Word or LibreOffice Writer, so that (probably) makes it the most userful tool I tried for docx tinkering in Python.

* `pip install pypdf2`: The PyPDF2 package works beautifully. Does what you want it to do, is fairly intuitive and allows loading, modifying and saving .pdf files. Nice tool all-round.

* `pip install docx2pdf`: It converts from docx 2 pdf, but only if you have MS Word installed. If not it, does not convert from docx 2 pdf.

* `pip install msoffice2pdf`: It converts from msoffice 2 pdf, but only if you have MS Word _or_ LibreOffice Writer installed, which is an improvement so that you do not need a $159.99 license that does not work on Linux.

* `pip install comtypes`: Its a general use file format conversor, and it also converts from .docx to .pdf... if you have MS Word _or_ LibreOffice Writer installed that is. In reallity, I think that msoffice2pdf uses comtypes underneath and it just offers a layer of abstraction for convinience.

* `pip install Spire.Doc`: The Spire.Doc package is proprietary software. It has a wonky free license that does not help you acomplish much other than trying out the package. Otherwise, it watermarks your modified .docx and .pdf files unless you pay for a license. It does not require the user to have installed any docx text processor, and its kinda intuitive to work with, but its still a paid proprietary software nonetheless, so not really a viable option for the vast majority of applications.


# Conclusions?

All in all, avoid .docx files like the plage, and try to move all your text processing and text operations to .pdf or simpler text format. If possible, convert your .docx files to .pdf and work with those, because if not you are likely to fall down the .docx rabbit hole of "why does the format of processed files and text changes _slightly_ from the original?" and the looming need of some .docx text processor.
