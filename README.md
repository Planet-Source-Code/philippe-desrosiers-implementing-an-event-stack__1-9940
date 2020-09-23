<div align="center">

## Implementing an event stack


</div>

### Description

Using DCOM? Remote instantiation? well now you can respond to events from a remote component without freezing the server app!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Philippe DesRosiers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/philippe-desrosiers.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/philippe-desrosiers-implementing-an-event-stack__1-9940/archive/master.zip)





### Source Code

```
<html>
<head>
<title>Implementing an event stack</title>
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:553679495 -2147483648 8 0 66047 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:none;
	page-break-after:avoid;
	mso-outline-level:1;
	mso-layout-grid-align:none;
	text-autospace:none;
	font-size:20.0pt;
	font-family:Tahoma;
	mso-bidi-font-family:"Times New Roman";
	mso-font-kerning:0pt;
	font-weight:bold;
	text-decoration:underline;
	text-underline:single;}
h2
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:none;
	page-break-after:avoid;
	mso-outline-level:2;
	mso-layout-grid-align:none;
	text-autospace:none;
	font-size:11.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;
	text-decoration:underline;
	text-underline:single;}
h3
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:3;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-bidi-font-family:"Times New Roman";
	font-weight:bold;
	text-decoration:underline;
	text-underline:single;}
p.MsoBodyTextIndent, li.MsoBodyTextIndent, div.MsoBodyTextIndent
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.2in;
	margin-bottom:.0001pt;
	mso-pagination:none;
	mso-layout-grid-align:none;
	text-autospace:none;
	font-size:10.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";
	color:blue;}
p
	{margin-right:0in;
	mso-margin-top-alt:auto;
	mso-margin-bottom-alt:auto;
	margin-left:0in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
 /* Page Definitions */
@page
	{mso-page-border-surround-header:no;
	mso-page-border-surround-footer:no;}
@page Section1
	{size:8.5in 11.0in;
	margin:.5in .5in .5in .5in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in;text-justify-trim:punctuation'>
<div class=Section1>
<h1>Implementing an Event Stack in VB</h1>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>With the advent of COM, DCOM and COM+, distributed
applications are fast becoming, indeed, have already become, a major focus for
new development tactics. It's just not enough anymore to write a puny little Access
database application and hope you won't need to implement it in a network
environment. More and more distributed applications are relying on an n-tier
model to get the job done. If you haven't yet had to deal with the increasing
demands of LAN, WAN and intranet-deployed apps, you might as well get ready,
because you'll have to eventually.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>DCOM allows VB programmers to create ActiveX servers that
can run as a standalone EXE on a remote machine. Because they run
out-of-process, unlike DLLs, multiple instances of the same class (perhaps
called by multiple applications on many different machines) can be accessed all
within the same process on the server machine. The EXE loads once, supplies the
class interface to whoever needs it, and when it's no longer needed, the EXE
unloads. Actually, it's much more complex than this, but if you want to learn
DCOM, read a book.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>So what's the problem?<u1:p></u1:p></h2>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>This remote instantiation is all well and good, a
revolution in computing, a watershed in distributed blah blah blah&#8230; with one
major drawback (at least for those ActiveX servers developed in VB). It's not
asynchronous. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>What does this mean? Asynchronous code is code that,
through multithreading or some other trickiness executes at the same time as
your application code. Instead of calling a method of your class to, say,
retrieve ten thousand records from an SQL database, then waiting while it
executes, then proceeding with your code, asynchronous execution would allow
you to call the method, which would return immediately, allowing you to
continue execution. In this example, when the class instance had completed
fetching the SQL result set, it would, say, raise an event to let your app know
that it was finished. If you use ADO with events, you know what I'm talking
about.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>The same system works in reverse. That is, when a control
or class raises an event trapped by the parent application, before the control
code can continue executing, the application has to execute its event code. For
example, when the Click event is raised, all the event code executes before
returning execution flow control to the ActiveX control. If you don't believe
me, try it yourself. You'll never receive a MouseUp event before you finish
processing the MouseDown event.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Now, normally, with an ActiveX control or DLL, all the
code runs in-process, i.e.: your app is the only one using it, so it doesn't
really matter if the control code stops execution while it waits for the event
to return. In fact, this is probably for the best. Who <i>wants</i> to receive the
MouseUp Event before you finish processing the MouseDown event? But with a DCOM
server component, running on a remote machine, that code can be executed by
hundreds of users at once, could be raising dozens of events at any one time,
from any number of classes. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>If the DCOM server component raises each of those events
to your app, then has to wait (while the application executes ten thousand
lines of code) before regaining execution flow control, the server is sitting
there, waiting for you to return flow control. While you're processing the
event code in your app, your preventing the server component from doing it's
job.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>This is, obviously, not a good thing.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>What's the answer to this dilemma? You got it, smart guy.
An event stack.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>What the heck is an event stack?<u1:p></u1:p></h2>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Oooo. A &quot;<i>Stack</i>&quot;. Scary word. Pointers.
Shades of linked lists and other murky memories from OOP theory courses you
slept through at school, right? Not so, my skittish friend. It's actually a
piece of cake to implement. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><b>Disclaimer: </b>Before you hard-core coders out there
start sending me e-mails, what we're doing here is not technically a stack.
Since a stack uses a Last In First Out (LIFO) implementation, it's unsuitable
for processing events in the order that they arrive (unless you're into that
sort of thing). Technically, I guess you could call this an event pipe, or
list, or funnel, but stack sounds cooler. Say it with me. <i>Stack.<o:p></o:p></i></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>An event stack exists for one purpose: to trap events and
store them for later processing. It's sort of like the transmission on your
car. Your transmission allows the engine and drive shaft to spin at two
different speeds without killing you and destroying your car. An event stack
allows your app to run and receive events from the remote DCOM component,
without taking execution flow away from that component. All right, so think of
your app as the wheels, and the DCOM component as the engine. Make more sense
now? No? Well I'm a VB geek, not a mechanic. Just stay tuned and you'll get it
eventually.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Okay, so what do I need?<u1:p></u1:p></h2>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Know anything about arrays? User-defined types? The VB
timer control? If so, you've got all the knowledge you need to implement an
event stack. If not, well... I guess you're out of luck.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Get to the point, already.<u1:p></u1:p></h2>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>What follows is the most basic way I know to create an
event stack. Obviously there are any number of improvements and changes you
could make. Among others: </p>
<p class=MsoNormal style='margin-left:.4in;text-indent:-.25in;mso-pagination:
none;mso-layout-grid-align:none;text-autospace:none'><span style='font-family:
Symbol'>·<span style='mso-tab-count:1'>      </span></span>Define a cEvent
class with a parameters collection instead of a UDT to hold your event
information.</p>
<p class=MsoNormal style='margin-left:.4in;text-indent:-.25in;mso-pagination:
none;mso-layout-grid-align:none;text-autospace:none'><span style='font-family:
Symbol'>·<span style='mso-tab-count:1'>      </span></span>Define an EventStack
collection with Push and Pop methods to contain the various events.</p>
<p class=MsoNormal style='margin-left:.4in;text-indent:-.25in;mso-pagination:
none;mso-layout-grid-align:none;text-autospace:none'><span style='font-family:
Symbol'>·<span style='mso-tab-count:1'>      </span></span>Use the SetTimer API
instead of the VB Timer control to trigger stack processing.</p>
<p class=MsoNormal style='margin-left:.4in;text-indent:-.25in;mso-pagination:
none;mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2><u1:p></u1:p>Step 1: Creating the event</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>First, we need to create a variable to hold
the information we're going to be receiving as parameters from the event. Let's
take a grossly simplified example. I've created an ActiveX DCOM component
that's running on a server machine. It exposes a class called <i>Xchat</i>,
whose purpose in life is to receive information via a <i>Post</i> method:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyTextIndent><span style='color:navy'>Public Sub</span> <span
style='color:windowtext'>Post</span> <span style='color:windowtext'>(</span><span
style='color:navy'>Optional </span><span style='color:windowtext'>Info1</span> <span
style='color:navy'>As Long, </span><span style='color:windowtext'>_</span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoBodyTextIndent><span style='color:navy'><span style="mso-spacerun:
yes">                 </span>Optional</span> <span style='color:windowtext'>Info2</span>
<span style='color:navy'>As String, </span><span style='color:windowtext'>_</span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoBodyTextIndent><span style='color:navy'><span style="mso-spacerun:
yes">                 </span>Optional</span> <span style='color:windowtext'>Info3</span>
<span style='color:navy'>As String</span><span style='color:windowtext'>)<u1:p></u1:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>And call an underlying function in a public
module which will pass this information to all the instances of the Xchat
class, by raising the <i>Dookie</i> event:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyTextIndent><u1:p></u1:p><span style='color:navy'>Public Event</span>
<span style='color:windowtext'>Dookie</span> <span style='color:windowtext'>(Info1</span>
<span style='color:navy'>As Long,</span> <span style='color:windowtext'>_</span></p>
<p class=MsoBodyTextIndent><span style='color:navy'><span style="mso-spacerun:
yes">                     </span></span><span style='color:windowtext'>Info2</span>
<span style='color:navy'>As String,</span> <span style='color:windowtext'>_</span></p>
<p class=MsoBodyTextIndent><span style='color:windowtext'><span
style="mso-spacerun: yes">                     </span>Info3</span> <span
style='color:navy'>As String</span><span style='color:windowtext'>)<u1:p></u1:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>This kind of DCOM server could be useful in hundreds
of ways; allowing a machine to poll the server to see how many connections are
active, as a component in a simple chat program, or a low-tech communications
protocol between apps running on different PCs. You get the idea.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>In this case we only want to trap one event, whose
parameters we know, so let's create a User-defined type (UDT) in a regular .BAS
module, to hold the event parameters.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>Public Type</span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New";color:blue'> </span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New"'>t_Event<u1:p></u1:p></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:blue'><span style="mso-spacerun: yes">   </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>FirstParam<span
style='color:blue'> </span><span style='color:navy'>As Long<u1:p></u1:p></span></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:blue'><span style="mso-spacerun: yes">   </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>SecondParam<span
style='color:blue'> </span><span style='color:navy'>As String<u1:p></u1:p></span></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:blue'><span style="mso-spacerun: yes">   </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>ThirdParam<span
style='color:blue'> </span><span style='color:navy'>As String<u1:p></u1:p></span></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>End Type<u1:p></u1:p><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>If you want to hold different events in your
stack (let's say a Timer event and an Error event), you might want to add an <i>EventID</i>
member to your UDT so the eventual processor of the events knows which event
it's processing. Likewise, you could add a <i>ControlID</i> if you want to trap
events from different controls, etc.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>&quot;But what if we don't know that all the events will
contain the same parameters?&quot;, I hear you ask. Good question. Like I said,
this is the <i>simplest</i> way to create an event stack. If you don't know
what parameters you'll be receiving, you could implement a <i>Parameters</i>
collection as a member of an <i>Event</i> class, which in turn would be
contained in an <i>EventStack</i> collection, etc., etc.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Easy enough, right? Wait. It gets even easier. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Step 2: Creating the stack</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Now we need to create a global variable in the same module
to hold all of our event information:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>Public</span><span style='mso-bidi-font-size:10.0pt;
font-family:"Courier New";color:blue'> </span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New"'>a_EventStack()<span style='color:blue'> </span><span
style='color:navy'>As</span><span style='color:blue'> </span>t_Event<u1:p></u1:p></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>This array will hold all of our miscellaneous event
information. We'll add events to the array, one at a time, as we receive them.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Step 3: Trapping the event</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Let's assume you're using the <i>Xchat</i> class in your
app. If you want to receive events from this component, you need to declare it
using the <i>WithEvents</i> keyword. So create a form in VB. In the code view for
the form, right after the <span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New"'>Option Explicit</span> statement (you <i>do</i> use Option
Explicit, don't you? Good.) type the following:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>Dim WithEvents</span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New";color:blue'> </span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New"'>xc_Remote<span style='color:blue'> </span><span
style='color:navy'>As</span><span style='color:blue'> </span>Xchat<u1:p></u1:p></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>Dim</span><span style='mso-bidi-font-size:10.0pt;
font-family:"Courier New";color:blue'> </span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New"'>b_LockStackProcessing<span style='color:blue'>
</span><span style='color:navy'>As Boolean<u1:p></u1:p></span></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>This, of course assumes you've correctly referenced the
remote component type libraries, configured it with <i>Dcomcnfg.exe</i>, and a
whole bunch of other stuff that is beyond the scope of this article.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Now, if you click on the object combo box at the top of
the code view window, you should see xc_Remote show up in the list of available
objects. Click on it, and like magic, we're transported to the <i>Dookie</i>
event.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Here is where the major advantage of an event stack starts
to become apparent. In the normal course of things, if this were a regular
class, a DLL, or an ActiveX control, you would do something like this:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>Private Sub</span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New";color:blue'> </span><span style='mso-bidi-font-size:
10.0pt;font-family:"Courier New"'>xc_Remote_Dookie (Info1<span
style='color:blue'> </span><span style='color:navy'>As Long,</span><span
style='color:blue'> </span>Info2<span style='color:blue'> </span><span
style='color:navy'>As String,</span><span style='color:blue'> </span>Info3<span
style='color:blue'> </span><span style='color:navy'>As String</span>)<u1:p></u1:p></span><span
style='color:blue'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'Execute
ten thousand lines of <o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">  
</span>'time-consuming, processor intensive code<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:green'><span style="mso-spacerun: yes">   </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'>End Sub<u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.2in;mso-pagination:none;mso-layout-grid-align:
none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>However, it's not. This component is raising dozens of
events per second, possibly to multiple clients, each of which wants to execute
its ten thousand lines of code before returning control to the server. See a
problem?</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>In order to get around this dilemma, we'll replace the
traditional event code with something like this:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">    </span><span style='color:navy'>Private Sub </span>xc_Remote_Dookie(Info1<span
style='color:navy'> As Long, </span>Info2<span style='color:navy'> As String, </span>Info3<span
style='color:navy'> As String</span>)<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Dim </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>l_Count<span
style='color:navy'> As Long<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>If </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>b_LockStackProcessing<span
style='color:navy'> Then<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">           </span>Exit Sub<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>End If<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>On Error GoTo </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>err_EmptyArray<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>l_Count<span
style='color:navy'> = UBound(</span>a_EventStack<span style='color:navy'>)<u1:p></u1:p></span></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>ReDim Preserve </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>a_EventStack(l_Count
+ 1)<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">    </span>err_Reentry:<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>On Error GoTo </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>0<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">       </span>a_EventStack(l_Count + 1).FirstParam =
Info1<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">       </span>a_EventStack(l_Count + 1).SecondParam =
Info2<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">       </span>a_EventStack(l_Count + 1).ThirdParam =
Info3<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Exit Sub<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">    </span>err_EmptyArray:<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">       </span>l_Count = 0<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>ReDim </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>a_EventStack(l_Count)<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Resume </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>err_Reentry<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">    </span>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>All this really does is grab the event
information and slap it into our event stack, and then returns control to the
component that raised the event. Since this code will execute in a fraction of
the time it would take to actually fully process the event, it doesn't take
control away from the server component for <i>too</i> long, and allows the
server to continue doing its job.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Step 4: Processing the event</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Now let's add a timer control <i>tmr_Event</i> to the
form, and set the interval property to some suitably small period, say 200 milliseconds.
The <i>Timer</i> event is where we'll process all the events we've trapped in
our stack, so we want to process the stack often enough to stay abreast of the
events being raised by the server, but not so often that we're constantly
interrupting client execution to handle the stack. After all, presumably this
application has other things to do.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Back into the code view for the form, let's add some code
to the <i>Timer</i> event:</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">    </span>Private Sub </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>tmr_Event_Timer<span
style='color:navy'>()<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Static </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>b_Reentry<span
style='color:navy'> As Boolean<u1:p></u1:p></span></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Dim </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>l_Count<span
style='color:navy'> As long<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Dim </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>i<span
style='color:navy'> As Integer<u1:p></u1:p></span></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>If </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>b_Reentry<span
style='color:navy'> Then<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">          
</span>Exit Sub<u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>End If<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>On Error Goto </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>err_EmptyArray<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>l_Count =<span
style='color:navy'> UBound(</span>a_EventStack<span style='color:navy'>)<u1:p></u1:p></span></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>On Error Goto </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>0<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">   </span><span style="mso-spacerun: yes"> </span><u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>b_Reentry =<span
style='color:navy'> True<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New";color:green'>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">      </span><span
style="mso-spacerun: yes"> </span>'.<u1:p></u1:p></span><span style='color:
green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'At this point, we
can execute some <o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'code to process
a_EventStack(0), since it is<u1:p></u1:p></span><span style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'the oldest event in
the stack. <u1:p></u1:p></span><span style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'.<u1:p></u1:p></span><span
style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:green'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">       </span>'Remove the oldest
event.</span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier
New"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
b_LockStackProcessing =<span style='color:navy'> True<u1:p></u1:p></span></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes"> </span><span style="mso-spacerun:
yes">  </span><span style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>If </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>l_Count = 0<span
style='color:navy'> Then<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">         
</span>Erase </span><span style='mso-bidi-font-size:10.0pt;font-family:"Courier
New"'>a_EventStack<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>Else<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">          </span>For </span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>i = 0<span
style='color:navy'> To </span>l_Count - 1<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">              </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>a_EventStack(i) =
a_EventStack(i + 1)<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">          </span>Next<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">         
</span>Redim Preserve </span><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New"'>a_EventStack(l_Count + 1)<u1:p></u1:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.75in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span>End If<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">       </span>b_LockStackProcessing =<span
style='color:navy'> False<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">   </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">    </span>err_EmptyArray:<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span></span><span
style='mso-bidi-font-size:10.0pt;font-family:"Courier New"'>b_Reentry =<span
style='color:navy'> False<u1:p></u1:p></span></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:"Courier New";
color:navy'><span style="mso-spacerun: yes">       </span><span
style="mso-spacerun: yes"> </span><u1:p></u1:p></span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;tab-stops:.25in;mso-layout-grid-align:
none;text-autospace:none'><span style='mso-bidi-font-size:10.0pt;font-family:
"Courier New";color:navy'><span style="mso-spacerun: yes">    </span>End Sub<u1:p></u1:p></span><span
style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><span style="mso-spacerun: yes">    </span></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>As you can see, we implement reentrancy protection on this
procedure with the <i>b_Reentry</i> variable, since the timer might tick ten
times before we finish processing the event stack, and we don't want to process
events out of turn. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><u1:p></u1:p>In addition to the standard reentrancy
protection, I've also added a global stack protection variable <i>b_LockStackProcessing</i>.
I use this so that no events can be added to stack while I'm resizing it. Since
the server component is running asynchronously to the application, it is
possible (though unlikely) to receive server events while resizing the stack. I
don't mind receiving events while I'm processing the stack, but I don't want to
overwrite existing events while I'm resizing, so I lock the stack. </p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>In this particular example, I only process one event off
the stack every time the timer ticks. Obviously if you plan on receiving more
than one event between timer ticks, you may want to process the entire stack
every time the timer ticks. Also, since the stack is an array, it takes some
time to shuffle all the events up one slot. You may want to implement a
collection instead, which, though it takes more memory and resources to handle,
can allow you to remove the event from the stack with one line of code.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><b>Note: </b>Another possible improvement on the code
here, which I leave you to implement, is a flexible timer. Every couple of
times you process the stack, you check to see if there are a large number of
events in the stack. If there are, you decrease the timer interval. If there
are very few events in the stack, you increase the timer interval. This means
that when the server is flipping out and flooding you with events, your app can
devote more time to handling those events, and when the server component is
twiddling its thumbs, your app can devote more processing time to other tasks.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>That's it?</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>That's all she wrote, boys and girls. See how easy that
was? While an event stack may not be necessary for small client-server
applications, it sure can save your bacon if you're deploying on an
Enterprise-wide scale. It can also be of enormous value if you want to dispatch
events asynchronously, just for the heck of it. Best of all, this technique can
be used in reverse, within your ActiveX server components, as a method stack,
processing method calls asynchronously so that the server component interface
is always available to clients.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Now go forth and multiply. Asynchronously.</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>Philippe DesRosiers</p>
<p class=MsoNormal style='mso-pagination:none;mso-layout-grid-align:none;
text-autospace:none'>e: philippe_desrosiers@karat.com<o:p></o:p></p>
</div>
<u1:p></u1:p>
</body>
</html>
```

