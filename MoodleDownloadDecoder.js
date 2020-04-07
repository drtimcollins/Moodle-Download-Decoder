var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var zipFile, csvFile;

// Find zip file and csv file
var d = fso.GetFolder('.');
var FileCollection = d.Files;
for(var f = new Enumerator(FileCollection); !f.atEnd(); f.moveNext()) {
	//WScript.Echo(f.item());
	if(fso.GetExtensionName(f.item()) == "zip")
		zipFile = f.item();
	if(fso.GetExtensionName(f.item()) == "csv")
		csvFile = f.item();		
}

//WScript.Echo("Zip file: " + zipFile);
//WScript.Echo("CSV file: " + csvFile);

// Unzip the archive
fso.CreateFolder('Submissions');
wsh.Run('tar -zxvf "' + zipFile + '" -C Submissions', 1, true); 

// Open excel to read csv data
var excel = new ActiveXObject("Excel.Application");
var oBook1 = excel.Workbooks.Open(csvFile);
var sheet = oBook1.Sheets.Item(1);
var N = sheet.UsedRange.Rows.Count;
// For each entry, copy file from Moodle-coded folder to a new MMU ID named folder
for(var n = 1; n < N; n++){
	var MoodleID = sheet.Range('B'+(n+1).toString()).Value;
	var MmuID = sheet.Range('D'+(n+1).toString()).Value.substring(0,8);
	fso.CreateFolder(fso.BuildPath('Submissions',MmuID));
	d = fso.GetFolder(fso.BuildPath(fso.BuildPath('Submissions',MoodleID),'File submissions'));
	FileCollection = d.Files;
	for(var f = new Enumerator(FileCollection); !f.atEnd(); f.moveNext()) {
		fso.MoveFile(f.item(), fso.BuildPath(fso.BuildPath('Submissions',MmuID),fso.GetFileName(f.item())));
	}
	fso.DeleteFolder(fso.BuildPath('Submissions',MoodleID));
//	WScript.Echo(MoodleID + " - " + MmuID);
}
excel.Quit();

WScript.Echo((N-1) + " Submissions decoded");
WScript.Quit(0);






