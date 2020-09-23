<div align="center">

## Using VB Compiler Directives


</div>

### Description

This Tutorial will allow you to use Visual Basic Compiler Directives such as

#IF..#Elseif..#Else..#End If

and

#Const

to create more robust and maintanable code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shawn Elliott](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shawn-elliott.md)
**Level**          |Beginner
**User Rating**    |4.5 (50 globes from 11 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shawn-elliott-using-vb-compiler-directives__1-22597/archive/master.zip)





### Source Code

```
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<title>Using VB Compiler Directives</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>Shawn Elliott</o:Author>
 <o:LastAuthor>Shawn Elliott</o:LastAuthor>
 <o:Revision>2</o:Revision>
 <o:TotalTime>41</o:TotalTime>
 <o:Created>2001-04-22T07:15:00Z</o:Created>
 <o:LastSaved>2001-04-22T07:15:00Z</o:LastSaved>
 <o:Pages>2</o:Pages>
 <o:Words>564</o:Words>
 <o:Characters>3218</o:Characters>
 <o:Company> </o:Company>
 <o:Lines>26</o:Lines>
 <o:Paragraphs>6</o:Paragraphs>
 <o:CharactersWithSpaces>3951</o:CharactersWithSpaces>
 <o:Version>9.2720</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
 <w:ActiveWritingStyle Lang="EN-US" VendorID="64" DLLVersion="131077"
  NLCheck="1">1</w:ActiveWritingStyle>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-font-kerning:0pt;}
h2
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:12.0pt;
	font-family:"Times New Roman";
	font-style:italic;}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:14.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	color:blue;}
p.MsoBodyTextIndent, li.MsoBodyTextIndent, div.MsoBodyTextIndent
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:1.5in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoBodyTextIndent2, li.MsoBodyTextIndent2, div.MsoBodyTextIndent2
	{margin:0in;
	margin-bottom:.0001pt;
	text-indent:.5in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoTitle align=left style='text-align:left'><span style='font-size:
16.0pt;mso-bidi-font-size:12.0pt'>Using VB Compiler Directives<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I don&#8217;t know how many times I have seen the following code</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;color:navy'>If</span><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'> DebugMode = true <span style='color:navy'>then</span><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'><span style='mso-tab-count:1'>                </span>Msgbox
&#8220;The Variable value is &#8220; &amp; SomeVar, vbokonly, &#8220;Debug&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;color:navy'>End if</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Or even the following</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><span
style='mso-tab-count:1'>                </span><span style='color:#339966'>&#8216;Uncomment
the following to debug this var<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
color:#339966'><span style='mso-tab-count:1'>                </span>&#8216;Msgbox
&#8220;The Variable value is &#8220; &amp; SomeVar, vbokonly, &#8220;Debug&#8221;</span><span
style='color:lime'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Many Visual Basic Programmers are not using one of the
powerful features of VB that equate it with other programming languages.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><b>Compiler Directives</b>.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:red'>&#8220;What are Compiler Directives?&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='text-indent:.5in'>Well, Compiler Directives are small
instructions that determine whether or not a piece of code will be included in
the Compile and Link process of creating an executable.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:red'>&#8220;What are the Compiler Directives to
use in VB?&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='text-indent:.5in'>In Visual basic you get the
following</p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:1'>            </span>#<span
style='color:navy'>Const</span></p>
<p class=MsoBodyTextIndent>This is private in the module it is defined.<span
style="mso-spacerun: yes">  </span>The Const items are NOT global to the
project only in their specific scope such as a form or class module</p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:1'>            </span>#<span
style='color:navy'>If</span></p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:2'>                        </span>This
is used to evaluate an expression of type #<span style='color:navy'>Const</span>
= n-expression </p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:1'>            </span>#<span
style='color:navy'>Elseif</span></p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:2'>                        </span>This
is used to evaluate an expression of type #<span style='color:navy'>Const</span>
= n-expression within and #<span style='color:navy'>IF</span> block</p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:1'>            </span>#<span
style='color:navy'>Else</span></p>
<p class=MsoBodyTextIndent2><span style='mso-tab-count:2'>                        </span>Code
within this sub-block is compiled if the #<span style='color:navy'>IF</span>
and #<span style='color:navy'>ELSEIF</span> blocks all evaluated to false</p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:1'>            </span>#<span
style='color:navy'>End If</span></p>
<p class=MsoNormal style='text-indent:.5in'><span style='mso-tab-count:2'>                        </span>This
ends the #IF Compiler Directive Block</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:red'>&#8220;Why would I want to use Compiler
Directives?<span style="mso-spacerun: yes">  </span>I have If Then Statements&#8221;<o:p></o:p></span></p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>If you
believe in adding additional code and using additional memory as well as CPU
cycles then Compiler Directives are not for you.<span style="mso-spacerun:
yes">  </span>Not to mention having Debug code or unwanted code in your final
exe simply because you forgot to comment one small line of code.<span
style="mso-spacerun: yes">  </span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h1>VB Specified Constants</h1>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>Visual
Basic defines some Compiler Constants automatically for you.<span
style="mso-spacerun: yes">  </span>These are:</p>
<p class=MsoNormal><span style='mso-tab-count:2'>                        </span>Win16<span
style='mso-tab-count:1'>  </span>&#8220;This indicates that the development
environment is 16-bit&#8221;</p>
<p class=MsoNormal><span style='mso-tab-count:2'>                        </span>Win32<span
style='mso-tab-count:1'>  </span>&#8220;This indicates that the development
environment is 32-bit&#8221;</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h1>How to use Compiler Directives</h1>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>Let&#8217;s take
a look at the first set of code we examined.<span style="mso-spacerun: yes"> 
</span>We notice it is a simple if then determining if the program needs to
show a Message Box with the value of a variable.<span style="mso-spacerun:
yes">  </span>How can we use a Compiler Directive to make this code more
efficient?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'>- THIS -<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;color:navy'>If</span><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'> DebugMode = true <span style='color:navy'>then</span><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'><span style='mso-tab-count:1'>                </span>Msgbox
&#8220;The Variable value is &#8220; &amp; SomeVar, vbokonly, &#8220;Debug&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;color:navy'>End if<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'>- CAN BE CHANGED TO -<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'>#<span style='color:navy'>Const</span> DebugMode =
true<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'>#<span style='color:navy'>If</span> DebugMode = true
<span style='color:navy'>then</span><o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'><span style='mso-tab-count:1'>                </span>Msgbox
&#8220;The Variable value is &#8220; &amp; SomeVar, vbokonly, &#8220;Debug&#8221;<o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:.5in'><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt'>#<span style='color:navy'>End If<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>Notice there was no real change in the code except we used
the #IF directive and ended the statement block with the #END IF statement to
check the value of a special Compiler Variable (which was defined with the
#CONST block)</p>
<p class=MsoNormal>These work the sameway as the If&#8230;Else&#8230;Elseif&#8230;End If
statement we are all used to except for one thing.<span style="mso-spacerun:
yes">  </span>If the condition being tested doesn&#8217;t prove true then the code
inside the block WILL NOT be included in the outputted code.<span
style="mso-spacerun: yes">  </span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h1>Other uses beside Debug</h1>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>The most
common use for Compiler Directives in other languages such as C and C++ are to
define sets of code for different operating systems and versions.<span
style="mso-spacerun: yes">  </span>We can do the same thing with visual basic.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
color:#339966'>&#8216;Programmer needs to determine what kind of code he is trying to
create by defining the<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
color:#339966'>&#8216;target OS Version here</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>Const</span> OSVersion = &#8220;Win9X&#8221;<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>If</span> OSVersion = &quot;Win9X&quot; <span
style='color:#333399'>Then</span><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><span
style='mso-tab-count:1'>                </span><span style='color:#339966'>&#8216;Programmer
needs to put specific Windows 95, 98 code here<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>ElseIf</span> OSVersion = &quot;WinNT&quot; <span
style='color:#333399'>Then</span><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><span
style='mso-tab-count:1'>                </span><span style='color:#339966'>&#8216;Programmer
needs to put specific Windows NT code here<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>ElseIf</span> OSVersion = &quot;Win2K&quot; <span
style='color:#333399'>Then</span><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><span
style='mso-tab-count:1'>                </span><span style='color:#339966'>&#8216;Programmer
needs to put specific Windows 2000 code here<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>Else</span><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'><span
style='mso-tab-count:1'>                </span><span style='color:#339966'>&#8216;Programmer
has not defined the OS Version<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>#<span
style='color:#333399'>End If<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
color:#333399'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
color:#333399'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h1>Final Notes</h1>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>An
important thing to remember is that #<span style='color:navy'>Const</span> and
Const variables cannot be interswitched.<span style="mso-spacerun: yes"> 
</span>If you try to use a #Const variable in place of a const or a const in
place of a #<span style='color:navy'>Const</span> variable VB will give you
Syntax errors.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>Also very
important is to Remember the scope of the #<span style='color:navy'>Const</span>
variable.<span style="mso-spacerun: yes">  </span>It is only within it&#8217;s module
like a Form, Module or Class Module.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h2>Shawn Elliott</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

