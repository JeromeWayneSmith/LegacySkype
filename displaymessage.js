// JavaScript Document displaymessage.js
var wsh = new ActiveXObject("WScript.Shell");
var timeout = 60;
wsh.popup( "Note: This message will disappear in " + timeout + " seconds. Just click button again to redisplay it.\n\n HELLO WORLD from SnapBack Apps!", timeout, "Hello World Sample Script with a Time-out" );

wsh = null;
