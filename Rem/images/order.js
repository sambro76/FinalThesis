function keyAction(){
var a;
a=window.parseInt(event.keyCode);
if(a==13) valid();
}
function resetbtn(){
document.f1.bcityo.disabled=true;
document.f1.bcityo.style.backgroundColor='#FFFFF7';
document.f1.bcountryo.disabled=true;
document.f1.bcountryo.style.backgroundColor='#FFFFF7';
document.f1.bbcityo.disabled=true;
document.f1.bbcityo.style.backgroundColor='#FFFFF7';
document.f1.bbcountryo.disabled=true;
document.f1.bbcountryo.style.backgroundColor='#FFFFF7';
document.f1.reset();
}

function valid(){
if(document.f1.D2.selectedIndex==0){document.f1.D2.focus(); window.scroll(100,0); alert("Please choose one type of CURRENCY you want to... ");}
else if(isNaN(document.f1.amnt.value)||document.f1.amnt.value<10|| document.f1.amnt.value=="") {document.f1.amnt.focus(); alert("Your amount field is not true, nor a valid integer number, \nPlease check your AMOUNT again... ");}

else if(document.f1.R3[0].checked==0&&document.f1.R3[1].checked==0){document.f1.R3[0].focus(); alert("Please choose an option for LOCAL BANK CHARGES... ");}
else if(document.f1.bfname.value==""|| !isNaN(document.f1.bfname.value)){document.f1.bfname.focus(); alert("Please check out your Beneficiary's Name carefully... ")}
else if(document.f1.blname.value==""|| !isNaN(document.f1.blname.value)){document.f1.blname.focus(); alert("Please check out your Beneficiary's Name carefully... ")}
else if(document.f1.bhb.value===""){document.f1.bhb.focus(); alert("Please check out your Beneficiary's Address carefully... ")}
else if(document.f1.bstreet.value==""){document.f1.bstreet.focus(); alert("Please check out your Beneficiary's Address carefully... ")}

else if(document.f1.D4.selectedIndex==0){document.f1.D4.focus(); alert("Please choose a Beneficiary's State/City... ");}
else if(document.f1.D4.selectedIndex==4&&document.f1.bcityo.value==""){document.f1.bcityo.focus(); alert("Please fill in the Beneficiary's State/City... ");}
else if(document.f1.D5.selectedIndex==0){document.f1.D5.focus(); alert("Please choose Beneficiary's Country... ");}
else if(document.f1.D5.selectedIndex==5&&document.f1.bcountryo.value==""){document.f1.bcountryo.focus(); alert("Please fill in the Beneficiary's Country... ");}

else if(document.f1.bbank.value==""){document.f1.bbank.focus(); alert("Please fill in Beneficiary's BANK... ");}
else if(document.f1.D5.selectedIndex==0){document.f1.D5.focus(); alert("Please choose the CITY of Beneficiary's Bank... ");}

else if(document.f1.D6.selectedIndex==0){document.f1.D6.focus(); alert("Please choose a State/City of Beneficiary's Bank... ");}
else if(document.f1.D6.selectedIndex==4&&document.f1.bbcityo.value==""){document.f1.bbcityo.focus(); alert("Please fill in the State/City of Beneficiary's Bank... ");}

else if(document.f1.D7.selectedIndex==0){document.f1.D7.focus(); alert("Please choose the COUNTRY of Beneficiary's Bank... ");}
else if(document.f1.D7.selectedIndex==5&&document.f1.bbcountryo.value==""){document.f1.bbcountryo.focus(); alert("Please choose the COUNTRY of Beneficiary's Bank... ");}

else if(document.f1.baccn.value==""){document.f1.baccn.focus(); alert("Please fill in the Account Number of Beneficiary... ");}
else if(document.f1.agree.checked==0) {document.f1.agree.focus(); alert("Please tick on the Agreement Conditions... ");}
else {document.f1.submit();}
}
