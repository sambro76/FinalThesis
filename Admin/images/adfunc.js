<!--
function actKey(){
	var a=parseInt(event.keyCode);
	if (a==13) {login();}
}
function Mod(x){return x-parseInt(x/16)*16}
function login(){
if (document.f1.lname.value=="" || !isNaN(document.f1.lname.value.charAt(0)) || (document.f1.lname.value.length<4) || document.f1.pwd.value=="" || (document.f1.pwd.value.length<3)){
	alert("Please fill in a certified account! ");
	document.f1.lname.value="";
	document.f1.pwd.value="";
	document.f1.lname.focus();
}
else {
	var c=parseInt(Math.random()*Math.pow(10,16));
	var d=parseInt(Math.random()*Math.pow(10,16));
	var arr=Array("a","b","c","d","e","f");
	var enc="";
	var end="";
	while(c>0){
		if(Mod(c)>9) {enc=arr[Mod(c)-10]+enc;}
		else {enc=String(Mod(c))+enc;}
		c=parseInt(c/16);
	}
	while(d>0){
		if(Mod(d)>9) {end=arr[Mod(d)-10]+end;}
		else {end=String(Mod(d))+end;}
		d=parseInt(d/16);
	}
	if(enc.length<16&&end.length<16) {
		for(i=0; i<=16-enc.length; i++) enc=enc + i;
		for(i=0; i<=16-end.length; i++) end=end + i;
	}
	document.f1.tmpid.value=enc + end;
	document.f1.action="adchid.asp";
	document.f1.submit();
   	document.f1.lname.value="";
	document.f1.pwd.value="";
	document.f1.tmpid.value="";
	document.f1.lname.focus();
}
}
function cls(){
	document.f1.lname.value="";
	document.f1.pwd.value="";
	document.f1.lname.focus();
}
// -->