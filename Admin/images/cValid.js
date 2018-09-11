<!--
function Mod(x){return x-parseInt(x/5)*5}
function chargeValid(){
for(var i=1; i<=(document.f1.elements.length-7); i++){
if(document.f1.elements[i].value==""||isNaN(document.f1.elements[i].value)||document.f1.elements[i].value.indexOf(" ")>=0){
	if(Mod(i-1)==0){
		if(document.f1.elements[i].value==""||isNaN(document.f1.elements[i].value)||document.f1.elements[i].value.indexOf(" ")>=0) {
			if(document.f1.elements[i].value.indexOf(" ")>=0){window.alert("Please do not use <<SPACE>> in each field... ");}
			document.f1.elements[i].focus();
			document.f1.elements[i].select();
			return 0;
			break;
		}
	}
	else {
		if(isNaN(document.f1.elements[i].value)||document.f1.elements[i].value.indexOf(" ")>=0) {
			if(document.f1.elements[i].value.indexOf(" ")>=0){window.alert("Please do not use <<SPACE>> in each field... ");}
			document.f1.elements[i].focus();
			document.f1.elements[i].select();
			return 0;
			break;
		}
		if(document.f1.elements[i].value==""){document.f1.elements[i].value="0";}
		if(i==document.f1.elements.length-7) {return 0;}
	}
}
}
}
// -->