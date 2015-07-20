// JavaScript Document displaymessage_execapp_runas.js
var wsh = new ActiveXObject("WScript.Shell");
var timeout = 60;
wsh.popup( "Note: This message will disappear in " + timeout + " seconds. Just click button again to redisplay it.\n\n This is a sample script, displaymessage_execapp_runas.js, that is included with this template. The execapp() function in jslib.js was called when you clicked the button. The name of this file was passed to the function as a parameter to wscript.exe. The \"runas\" verb was used to elevate the execution.\n\n Hook up your App's custom function to a button or use functions in the jslib.js file!", timeout, "AppTemplate Sample Script with a Time-out" );

wsh = null;
