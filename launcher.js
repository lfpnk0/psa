//
//Set global variables
//
var fso = new ActiveXObject("Scripting.FileSystemObject");
var aShell = new ActiveXObject("Shell.Application");
var wShell = new ActiveXObject("Wscript.Shell");
var network = new ActiveXObject("WScript.Network");
//
var desktop = wShell.SpecialFolders(4);
var psaDir = fso.GetSpecialFolder(2) + '\\psa';
var appName = document.getElementById('launcher').applicationname;
var appDir = psaDir + '\\' + appName;
var appPath = psaDir + '\\' + appName + '\\app.hta';
var sysUrl = document.getElementsByTagName('script')[0].src.replace("launcher.js", "");
var appUrl = sysUrl + appName;
var zipUrl = '';
var confPath = appDir + '\\conf.js';
var confUrl = appUrl + '/conf.js';
var uname = network.UserName;
var profUrl = appUrl + '/profiles/' + uname + '.js';
var oldVer = '0.0.0';
var newVer = '0.0.0';
var offerUpg = 0;
//
//get local and most recent app versions
//
if(fso.FileExists(confPath)){
	var confObj = fso.OpenTextFile(confPath, 1);
	var t = confObj.ReadAll();
	oldVer = t.substring(t.indexOf('"version":')+10, t.indexOf(';', t.indexOf('"version":')));
}
var xmlHttp = new XMLHttpRequest();
xmlHttp.onreadystatechange = function() {
	if (this.readyState == 4 && this.status == 200) {
		t = xmlHttp.responseText;
		newVer = t.substring(t.indexOf('"version":')+10, t.indexOf(';', t.indexOf('"version":')));
		zipUrl = appUrl + '/' + newVer + '.zip';
		checkUpg();
	}
};
xmlHttp.open('GET', confUrl, true);
xmlHttp.send();
//
// Upgrade app files if needed
//
function checkUpg(){
	var ov = oldVer.split(".");
	if(ov.length < 3){ov.push(0); ov.push(0); ov.push(0);}
	var nv = newVer.split(".");
	if(nv.length < 3){nv.push(0); nv.push(0); nv.push(0);}
	for (i = 0; i < 3; i++) {
		if(ov[i]+1 < nv[i]+1){offerUpg = 1;}
	}
	if(offerUpg > 0){
		var upg = confirm('Update Application Files?');
	}
	if(upg){
		if(!fso.FolderExists(psaDir)){fso.CreateFolder(psaDir);}
		if(!fso.FolderExists(appDir)){fso.CreateFolder(appDir);}
		fso.DeleteFile(appDir+'\\*', true);
		fso.DeleteFolder(appDir+'\\*', true);
		fso.CopyFile(zipUrl, appDir+'\\');
		var inZip=aShell.NameSpace(appDir+ '/' + newVer + '.zip').items;
		aShell.NameSpace(appDir + '/').CopyHere(inZip);
		fso.CopyFile(profUrl, appDir+'\\profiles\\');
		fso.CopyFile(appPath, desktop+'\\');
	}
	startApp();
}
//
// Start Application
//
function startApp(){
	if(fso.FileExists(appPath)){
		aShell.Run(appPath);
		window.close;
	}
}