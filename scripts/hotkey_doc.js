//add hotkey klee v1.0 ;-)
var lcKeyArray = new Array('!','"','£','@');
var lcLocationArray = new Array('../../ciphers/default.htm','../../spies/default.htm','../../codemaster/default.asp','../../default.htm');

function doHotKey(lc){
	for(i=0; i < lcKeyArray.length; i++){
		if(document.layers){ // NN4+ ;-)
			if (lcKeyArray[i] == String.fromCharCode(lc.which)){
				window.location = lcLocationArray[i];
			}
		} else if(document.all){ // IE4+ ;->
			if(lcKeyArray[i] == String.fromCharCode(event.keyCode)){
				window.location = lcLocationArray[i];
			}
		} else if(document.getElementById){ // NN6 :8)
			if(lcKeyArray[i] == String.fromCharCode(lc.which)){
				window.location = lcLocationArray[i];
			}
		}
	}
}
if(document.layers){
	document.captureEvents(Event.KEYPRESS);
}
document.onkeypress=doHotKey;