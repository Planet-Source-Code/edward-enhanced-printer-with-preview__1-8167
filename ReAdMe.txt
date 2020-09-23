Project:	qbd Enhanced Printer Control with Preview
Copyright:	© 2000 qbd software ltd
Author:		edward moth

=====================================================================

PLEASE READ THE BORING WARRANTY AND LICENSE INFORMATION - FAILURE TO
DO SO WILL MAKE YOU LIABLE TO BE TUTTED AT UNDER THE PROVISIONS OF 
SECTION 131.9(B) OF THE PROTECTION OF edward moth ACT (1999) - HM 
GOVERNMENT (UNITED KINGDOM).  PLEASE ALSO NOTE THAT UNDER SECTION 
112.3 OF SAID ACT, ALL PAYMENTS TO EDWARD SHOULD BE MADE IN CASH IN BROWN ENVELOPES - BET NO ONE READS THIS - HEHEHE - THEY NEVER DO.

=====================================================================

A.	PURPOSE
B.	REQUIREMENTS
C.	INSTRUCTIONS FOR SETTING UP AND USE (PROPERTIES/METHODS)
D.	WARRANTY AND LICENSE
E.	CONTACT

=====================================================================

A PURPOSE:

Improved Print handling with Preview.  Documents are constructed by
adding TextItems to the qcPrinter Class.  Each TextItem can have it's
own font attributes and alignment (Left, Right, Centre or Full 
Justification).  Each item can have left and right indents.

This is a WIP (work in progress) - further development may consider
including support for pictures, multiple font styles and attributes
(eg. basic RTF support), direct positioning, tables and column
printing.

=====================================================================

B REQUIREMENTS:

Just the basics

=====================================================================

C INSTRUCTIONS FOR USE:

WARNING:  The project has limited error handling support.  Assumption
is made that the programmer will not use stupid property value - e.g.
margins that are larger than the page ... duh.

A Test Project Group is included that shows some of the uses of the
component (see frmTest in the qPrinterTest project)

To use the class(es) in another project, add the qPrinter project to
the project group.  Add a reference from you project. Alternatively, 
compile the qPrinter project and add a reference to qPrinter.dll.

PROPERTIES, METHODS and EVENTS

PROPERTIES:

	ItemCount	(Read/Long) Number of TextItems held

	MarginBottom 	(Single) Margin values - displayed in the
	MarginLeft	current Scalemode (see below)
	MarginRight
	MarginTop

	Pages		(Read/Integer) Total number of pages in the
			document
	PageSize	(Enum) A3, A4, A5 or B4
	ScaleMode	(Enum) Twip, Inch, Centimetre or Millimetre
	
	TextItem(IndexKey)
		.Alignment	(Enum) Left, Centre, Right, Justify

		.FontName	(String) Name of font
		.FontSize	(Single) Size of font
		.FontBold, .FontItalic, .FontUnderline (Boolean)
		.FontColor	(Long) Colour of font

		.IndentLeft	(Single) TextItem distance from left
		.IndentRight	and right margins

		.Text		(String)

METHODS:

	AddText		Add a Text Item, using the properties above
			(a Key may be specified for future reference)
	Preview		Display the Preview form with the current
			document
	PrintDoc	(StartPage, EndPage) Print the specified
			pages of the current document
	RemoveItem	(IndexKey) Remove the specified Item, Returns
			True if successful, False if not.
	ResetItems	Clears all the TextItems in the current
			document

EVENTS: None

	
=====================================================================

D BLAH ... THE BORING BITS

WARRANTIES:
All code is provided 'as is', without warranties of any kind whatsoever
no matter who you say your dad is, even if it is expressed or implied
(the warranties that is, not your dad).

LIABILTY:
qbd software ltd and edward moth accept no liability whatsoever even
if you are an Arsenal supporter.  By using this code you accept
that Manchester United are the greatest football team of all time and
that Doncaster Belles was robbed in the Women's FA Cup Final
(okay ... if you're not happy with that, I'll let you off).

LICENSE:
You are free to use and modify any of the code but it would be nice
if you left references to qbd software limited, qbd and edward moth in
place :-) obviously if you nick it and make out it's your own coding
the management and employees of qbd software ltd will be forced to
take serious retaliatory action such as calling you a southern wuss.

You can freely distribute the zip and the code but give us credit
(Amex Platinum for preference but 'Thanx to edward moth and qbd
software ltd' would suffice).  If you wish to contant us then please
see the instructions below.

=====================================================================

E CONTACT US:

If you wish to contact edward moth or qbd software then
please follow these instructions:

Categorise your mail:

Mail Type One:	
		* Begging Letters
		* Spam
		* Complaint
		* Abusive
		* Marriage proposal
		* Advertising
		* Porn related
		* Pyramid selling
		* Stupid
		* Contains a virus (we particularly dislike common colds)
		* Contains blank lines because you sent it before you
		  had written it
		* To brag about how much better you are a PSX games
		  than edward (highly unlikely)
		
	or - 	* If you have any contagious medical condition
		* You have an obsession that encompasses your whole life
		* You are William Hague MP

	Send your Mail to: 	trash@qbdsoftware.co.uk
				It will be treated with the utmost
				respect and dealt with promptly

Mail Type Two:

		* Praise
		* Contracts for Tender
		* Job offers (minimum salaries apply - make sure it's
		  good or we'll laugh)
		* Funny
		* Worth reading
		* Exciting Ideas that will make us rich (and that does
		  not include selling perfume)

	or:	* You are Bill Gates regarding 'those share options'
		* You are edward's mom
			

	Send your Mail to:	edward@qbdsoftware.co.uk


