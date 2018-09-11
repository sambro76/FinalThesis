function click(e) {
if (document.all) {
if (event.button==2||event.button==3) {window.alert("Copyright ©2004 by c_samnang@yahoo.com ");return false;}
}
if (document.layers) {
if (e.which == 3) {window.alert("Copyright ©2004 by c_samnang@yahoo.com ");return false;}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;
