////////////////////////////////////////////////////////
// (c)2009 CodeCentrix Software. All rights reserved. //
////////////////////////////////////////////////////////


if (WScript.Arguments.length == 0)
{
	// Explanatory welcome message.
	var text = "Twebst Library Quick Tour\n\n";

	text += "This is a demonstration of some common scenarios of using Twebst Library.\n";
	text += "The script will run automatically wihtout the need for user input.\n\n";
	text += "1. Test a web application by generating input and extracting expected results.\n";
	text += "2. Extract data from web pages.\n";
	text += "3. Automate web actions and data-entry in web pages.\n";
	text += "4. Get access to native DOM and dynamically modify HTML documents.\n";

	WScript.Echo(text);
}

// Create a Twebst core object.
var core = null;

try
{
	core = new ActiveXObject("Twebst.Core");
}
catch (ex)
{
	//  It could be Win64.
	if (WScript.Arguments.length == 0) // First normal try.
	{
		var fileSystem = WScript.CreateObject("Scripting.FileSystemObject");
		var shell      = WScript.CreateObject("WScript.Shell");
		var wsPath     = shell.ExpandEnvironmentStrings("%windir%") + "\\SysWOW64\\wscript.exe";
		
		if (fileSystem.FileExists(wsPath))
		{
			// Launch QuickTour.js with 32 bit version of wscript.exe
			// Add "1" parameter to command line to know if it is the second attempt.
			var qtCmd  = "\"" + wsPath + "\" \"" + WScript.ScriptFullName + "\" 1";
			shell.Run(qtCmd);
		}
		else
		{
			WScript.Echo("Error: can not create Twebst.Core object!");
		}
	}
	else
	{
		// Error creating Core object on second attempt.
		WScript.Echo("Error: can not create Twebst.Core object!");
	}

	WScript.Quit(1);
}

// Start a new browser and navigate to the quick tour web page.
var browser = core.StartBrowser("http://www.codecentrix.com/tests/qt1.htm");
browser.ShowBrowserWindow(browser.SHOW_MAXIMIZE);
browser.ShowBrowserWindow(browser.SHOW_IN_FOREGROUND);

// Script initialization.
var elem       = null;
var SLEEP_TIME = 100;

try
{
	core.useHardwareInputEvents = true;
}
catch (e)
{
	// useHardwareInputEvents is NOT available in FREE version.
}


// Find the bulleted list HTML native element.
var listElem = browser.FindElement("ul").nativeElement;

// Find the edit control in the calculator application.
var editElem   = browser.FindElement("input text");
var nativeEdit = editElem.nativeElement;

// STEP ONE: Perform a test on "+" operator.
listElem.innerHTML += "<li>Step one: &nbsp;&nbsp;test the \"+\" operator. Perform 1 + 2</li>";
WScript.Sleep(SLEEP_TIME);

// Press 1 + 2 = calculator buttons.
elem = browser.FindElement("input button", "text=1");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text=+")
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text=2");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text==");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

// STEP TWO: Verify the result.
editElem.Highlight();
if (nativeEdit.value == 3)
{
	listElem.innerHTML += "<li>Step two:&nbsp;&nbsp; the result is 3. Test passed!</li>";
}
else
{
	listElem.innerHTML += "<li>Step two: &nbsp;&nbsp;the result is NOT 3. Test failed!</li>";
}
WScript.Sleep(SLEEP_TIME);

// STEP THREE: Reset the calculator by pressing "C".
listElem.innerHTML += "<li>Step three: reset the calculator by pressing \"C\".</li>";
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text=C");
elem.Highlight();
elem.Click();


// STEP FOUR: Perform a test on "-" operator.
listElem.innerHTML += "<li>Step four: &nbsp;test the \"-\" operator. Perform 5 - 2</li>";
WScript.Sleep(SLEEP_TIME);

// Press 5 - 2 = calculator buttons.
elem = browser.FindElement("input button", "text=5");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text=-");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text=2");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

elem = browser.FindElement("input button", "text==");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);


// STEP FIVE: Verify the result.
editElem.Highlight();
if (nativeEdit.value == 3)
{
	listElem.innerHTML += "<li>Step five: &nbsp;the result is \"3\". Test passed!</li>";
}
else
{
	listElem.innerHTML += "<li>Step five: &nbsp;the result is <font color=\"red\">NOT</font> \"3\". Test failed!</li>";
}
WScript.Sleep(SLEEP_TIME);

// Insert "Next" link.
var nextCell = browser.FindElement("td", "id=next_step");
nextCell.nativeElement.innerHTML += "<a href=\"./qt2.htm\"><b>Next &gt&gt</b></a>";
WScript.Sleep(SLEEP_TIME);

try
{
	// Go to the next step of this "Quick tour".
	var nextLink = core.AttachToNativeElement(nextCell.nativeElement.children(0));
	nextLink.Highlight();
	
	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}
catch (e)
{
	// AttachToNativeElement is NOT available in FREE version.
	var nextLink = browser.FindElement("a", "text=Next >>");
	nextLink.Highlight();

	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}



/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// Step TWO ///////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Script initialization.
var fileName    = "";
var textFile    = null;
var excelApp    = null;
SLEEP_TIME  = 1200;

// Can not start Excel application. Create a new file.
var shellObj    = WScript.CreateObject("WScript.Shell");
var myDocuments = shellObj.SpecialFolders("MyDocuments");

try
{
	// Initialize the Excel document.
	excelApp = new ActiveXObject("Excel.Application");
	excelApp.Workbooks.Add;
}
catch (except)
{
	fileName    = myDocuments + "\\TwebstExchangeRates.txt";
	textFile    = WScript.CreateObject("Scripting.FileSystemObject").CreateTextFile(fileName , true);
}


try
{
	core.useHardwareInputEvents = true;
}
catch (e)
{
	// useHardwareInputEvents is NOT available in FREE version.
}

// Find the bulleted list HTML native element.
var listElem = browser.FindElement("ul").nativeElement;

// STEP ONE: Find the exchange rates table.
listElem.innerHTML += "<li>Step one: &nbsp;find the table to be extracted.</li>";
WScript.Sleep(SLEEP_TIME);

var table = browser.FindElement("table", "id=exchange_rates");
//var nativeTable = table.nativeElement;
table.Highlight();


// STEP TWO: Iterate through table cells and save data in excel or text format.
listElem.innerHTML += "<li>Step two: &nbsp;Iterate through each cell in the table and extract the data.</li>";
WScript.Sleep(SLEEP_TIME);


function Spaces(n)
{
	var s = "";
	for (i = 0; i < n; ++i)
	{
		s += " ";
	}

	return s;
}


function SaveCell(excelApp, row, col, currentCol)
{
	var cell = excelApp.ActiveSheet.Cells(row + 2, col + 2);

	cell.Formula             = currentCol.nativeElement.innerText;
	cell.Interior.ColorIndex = 24;

	if (row != 0)
	{
		cell.HorizontalAlignment = 4;
		cell.Borders.LineStyle   = 1;

		// Insert the image in the excel document.
		core.searchTimeout = 0;
		var img = currentCol.FindElement("img");
		core.searchTimeout = 10000;

		if (img != null)
		{
			var imgFileName = myDocuments + "\\" + row + "_" + col + ".jpg";
			img.SaveElementImage(imgFileName);

			cell.Select();
			excelApp.ActiveSheet.Pictures.Insert(imgFileName);
		}
	}
	else
	{
		for (i = 0; i < 5; ++i)
		{
			var emptyCell = excelApp.ActiveSheet.Cells(2, i + 3);
			emptyCell.Interior.ColorIndex = 24;
		}
	}
}


var rows = table.FindAllElements("tr");
for (r = 0; r < rows.length; ++r)
{
	var currentRow = rows(r);
	var cells      = currentRow.FindAllElements("td");

	for (c = 0; c < cells.length; ++c)
	{
		var currentCol    = cells(c);
		var oldBackground = currentCol.nativeElement.style.background;

		currentCol.nativeElement.style.background = "red";
		WScript.Sleep(50);
		currentCol.nativeElement.style.background = oldBackground;

		// Extract data from the current cell.
		if (excelApp != null)
		{
			SaveCell(excelApp, r, c, currentCol);
		}
		else if (textFile != null)
		{
			// Extract the data in TXT format.
			var textToWrite = currentCol.nativeElement.innerText;
			if (textToWrite.length > 0)
			{
				textToWrite = Spaces(10 - textToWrite.length) + textToWrite;
				textFile.Write(textToWrite + "\t");
			}
		}
	}

	if (textFile != null)
	{
		textFile.Write("\r\n");
	}
}

if (textFile != null)
{
	textFile.Close();
}

if (excelApp != null)
{
	// STEP THREE: .
	listElem.innerHTML += "<li>Step three: Show the extracted data in Excel format.</li>";
	WScript.Sleep(SLEEP_TIME);

	excelApp.ActiveSheet.Cells(10, 1).Select();
	excelApp.Visible = true;
}
else
{
	// STEP THREE: .
	listElem.innerHTML += "<li>Step three: Show the extracted data in TXT format (Excel application couldn't be found).</li>";
	WScript.Sleep(SLEEP_TIME);

	shellObj.Run("\"" + fileName + "\"");
	WScript.Sleep(SLEEP_TIME);
}


// Insert "Next" link.
var nextCell = browser.FindElement("td", "id=next_step");
nextCell.nativeElement.innerHTML += "<a href=\"./qt3.htm\"><b>Next &gt&gt</b></a>";
WScript.Sleep(SLEEP_TIME);


try
{
	// Go to the next step of this "Quick tour".
	var nextLink = core.AttachToNativeElement(nextCell.nativeElement.children(0));
	nextLink.Highlight();
	
	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}
catch (e)
{
	// AttachToNativeElement is NOT available in FREE version.
	var nextLink = browser.FindElement("a", "text=Next >>");
	nextLink.Highlight();

	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}



/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// Step THREE /////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Script initialization.
SLEEP_TIME = 1200;
try
{
	core.useHardwareInputEvents = true;
}
catch (e)
{
	// useHardwareInputEvents is NOT available in FREE version.
}


WScript.Sleep(SLEEP_TIME);
var elem = browser.FindElement("input text", "text=First name:");
elem.Highlight();
elem.InputText("D.");

elem = browser.FindElement("input text", "text=Last name:");
elem.Highlight();
elem.InputText("Adrian");

elem = browser.FindElement("select", "text=Gender:");
elem.Highlight();
elem.Select("male");

elem = browser.FindElement("input password", "text=Password:");
elem.Highlight();
elem.InputText("mypassword");

elem = browser.FindElement("input password", "text=Re-type password:");
elem.Highlight();
elem.InputText("mypassword");

elem = browser.FindElement("select", "text=Country:");
elem.Highlight();
elem.Select("Romania");

elem = browser.FindElement("input file", "text=Your photo:");
elem.Highlight();
elem.InputText(WScript.ScriptFullName);

elem = browser.FindElement("input radio", "text=Between 21 and 30");
elem.Highlight();
elem.Click();

elem = browser.FindElement("input checkbox", "text=Just a checkbox");
elem.Highlight();
elem.Click();

elem = browser.FindElement("input button", "text=Don't Submit!");
elem.Highlight();
elem.Click();
WScript.Sleep(SLEEP_TIME);

// Insert "Next" link.
var nextCell = browser.FindElement("td", "id=next_step");
nextCell.nativeElement.innerHTML += "<a href=\"http://www.yahoo.com\"><b>Next &gt&gt</b></a>";
WScript.Sleep(SLEEP_TIME);


try
{
	// Go to the next step of this "Quick tour".
	var nextLink = core.AttachToNativeElement(nextCell.nativeElement.children(0));
	nextLink.Highlight();
	
	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}
catch (e)
{
	// AttachToNativeElement is NOT available in FREE version.
	var nextLink = browser.FindElement("a", "text=Next >>");
	nextLink.Highlight();

	WScript.Sleep(SLEEP_TIME);
	nextLink.Click();
}



/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// Step FOUR //////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Script initialization.
try
{
	core.useHardwareInputEvents = true;
}
catch (e)
{
	// useHardwareInputEvents is NOT available in FREE version.
}

// Count the number of links in the web page.
var numberOfAnchors = browser.FindAllElements("a").length;

// Count the number of images.
var numberOfImages = browser.FindAllElements("img").length;

// Count the number of tables.
var numberOfTables = browser.FindAllElements("table").length;

// Count the number of inputs.
var numberOfInputs = browser.FindAllElements("input text").length;

// Get the HTML document and the <body> element.
var document = browser.topFrame.document;
var body     = document.body;

WScript.Sleep(750);

// Modify the web site.
var htmlToAdd = '<br><br><h1 style="text-align: left;">&nbsp;<i>Twebst</i> quick tour (4 of 4)</h1><br>' + 
                '<h3 style="text-align: left;"><font color="green">&nbsp;&nbsp;<b>' + 
                'Web data extraction and dynamically modifying web documents.</b></font></h3>' +
                '<br><br>';
htmlToAdd += '<div style="text-align: left;">&nbsp;&nbsp;<b>Statistics:</b><div>';
htmlToAdd += '<div style="text-align: left;">&nbsp;&nbsp;&nbsp;&nbsp;Links: '  + numberOfAnchors + '</div>';
htmlToAdd += '<div style="text-align: left;">&nbsp;&nbsp;&nbsp;&nbsp;Images: ' + numberOfImages  + '</div>';
htmlToAdd += '<div style="text-align: left;">&nbsp;&nbsp;&nbsp;&nbsp;Tables: ' + numberOfTables  + '</div>';
htmlToAdd += '<div style="text-align: left;">&nbsp;&nbsp;&nbsp;&nbsp;Edit-boxes: ' + numberOfInputs  + '</div>';
htmlToAdd += '<div style="text-align: center;">Thank you for your time!</div>';
htmlToAdd += '<center><b><a style="color: green" href="http://www.codecentrix.com">&gt;&gt;Go back to Codecentrix web site&lt;&lt;</a></b></center><hr><br>';

body.innerHTML = htmlToAdd + body.innerHTML;
