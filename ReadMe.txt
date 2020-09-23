Basic Stock Charting Project V1.0 by John Sugas 2000-2002 jsugas@mei.net. 
No warranties or guarantees of any kind are included. I'm sharing this for personal use only, 
don't sell this code... Any code borrowed from other programmers has been commented
as such, as well as a module with nothing but said code. Unconventional... maybe, but
Their efforts deserved a special place in this project... Thanks to all who contributed.

This charting program won't put MetaStock or TradeStation out of business but it 
does show the basic principles of charting stock data using VB. The code may have 
bugs and non-optimized code sections.... I certainly could have implemented some of the 
functionality better. Comments are not included on every line only where needed to 
record the more complex statements... (Now what did I do that for again??? <s>). 
This was coded up on an NT box and works fine on a W2K box also..... 
But I have no idea if it will work Ok on W95-W98.... You're on your own there folks.
I have a feeling some of the API routines may be problems for the W95-W98 crowd....
A TypeLibrary has been made for this project. Source code is included. It is ANSII
so that will work on the W95-98 boxes, but who knows....

Data can be intraday(minutes-hours) or End of Day(EOD). Only comma separated data
is utilized currently. You could change that if needed. The first line of stock data 
files sometimes specifies the item order of the data. The loader looks for a couple types.
More could be added to the Select Case if needed. The symbol is found by looking at 
the file name, example -> QQQ~1m-5-13-02.dat  The symbol info is the first part of the
file name and ~ is the end token. Optionally the symbol panel on the status bar can be 
DblClk'ed when it reads "???" for a manual edit of the symbol. Holidays and missing 
data are not taken into consideration nor is data updated in real-time... Sorry... 
Something for you to do <s>. A couple of data files are included.

Chart mouse interaction includes left and right click functionality. The R.Clk gives you 
the standard pop-up menu for options and cancels. The L.Clk does several things depending 
on location. In the price plot panel a MouseDn will bring up the crosshairs. The time of 
this event is saved. If a second click is made before the time expires, a data window will 
be shown in the upper right corner of the price panel. Holding the L.Button down and 
moving around will give the data info for the different price bars. L.Button will also 
adjust the sizes of the panels, move the mouse to a divider line and drag to new position.
The default values for the divider locations may be totally out of range. I wrote the code
on a 17" monitor then when testing on a 21" one divider was locked. I added a test to 
check for that kind of condition and correct it before the chart gets displayed but
can't foresee every situation. If you find that problem go to the options dialog
and manually change the locations. Defaults in the MakeINI function can be changed also.

Toolbar functions include opening a file, barspacing, redraw, options, screen capture,
basic drawing tools, indicator functions. Something to note; the drawing tools functions 
are utilizing a modal loop instead of the MouseMove event. Use the R.Clk Cancel Drawing 
entry to excape. Drawing tools are not saved... maybe V2.0....

Scrolling functions can be done by Toolbar or keyboard. By Toolbar, the Left mouse button 
scrolls left 1 bar and the right button scrolls right 1 bar. If shift is held down when 
clicking the scroll amount is 10 or greater, the setting is configurable in the options
dialog. Using the keyboard, CTRL must be held down for all scroll ops. The Right & Left 
arrow keys will scroll 1 bar, the PageUp & PageDn keys will scroll the ScrollAmount,
Ctrl-Home will send you to the beginning of the chart and Ctrl-End will take you back
to the very end (most recent data).


Future enhancement possibilities:

MDI interface
Scaling options for price & indicators
Save drawing tools to file
Add more common stock indicators
Better indicator management and saving of parameters 
	(indicator collection with binary storage file)
User defined indicator functions
Multiple time frame plots on same chart
Overlay price with second data series
...