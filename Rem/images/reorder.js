function keyAction(){
var a;
a=window.parseInt(event.keyCode);
if(a==13) valid();
}

function valid(){
if(document.f1.D2.selectedIndex==0){document.f1.D2.focus(); window.scroll(100,0); alert("Please choose one type of CURRENCY you want to... ");}
else if(isNaN(document.f1.amnt.value)||document.f1.amnt.value<10|| document.f1.amnt.value=="") {document.f1.amnt.focus(); alert("Your amount field is not true, nor a valid integer number, \nPlease check your AMOUNT again... ");}

else if(document.f1.R3[0].checked==0&&document.f1.R3[1].checked==0){document.f1.R3[0].focus(); alert("Please choose an option for LOCAL BANK CHARGES... ");}
else if(document.f1.bfname.value==""|| !isNaN(document.f1.bfname.value)){document.f1.bfname.focus(); alert("Please check out your Beneficiary's Name carefully... ")}
else if(document.f1.blname.value==""|| !isNaN(document.f1.blname.value)){document.f1.blname.focus(); alert("Please check out your Beneficiary's Name carefully... ")}
else if(document.f1.bhb.value===""){document.f1.bhb.focus(); alert("Please check out your Beneficiary's Address carefully... ")}
else if(document.f1.bstreet.value==""){document.f1.bstreet.focus(); alert("Please check out your Beneficiary's Address carefully... ")}
else if(document.f1.bcity.value==""){document.f1.bcity.focus(); alert("Please check out your Beneficiary's CITY carefully... ")}
else if(document.f1.bcountry.value==""){document.f1.bcountry.focus(); alert("Please check out your Beneficiary's COUNTRY carefully... ")}

else if(document.f1.bbank.value==""){document.f1.bbank.focus(); alert("Please fill in Beneficiary's BANK... ");}
else if(document.f1.bbcity.value==""){document.f1.bbcity.focus(); alert("Please fill in BANK Location... ");}
else if(document.f1.bbcountry.value==""){document.f1.bbcountry.focus(); alert("Please fill in BANK Location... ");}

else if(document.f1.baccn.value==""){document.f1.baccn.focus(); alert("Please fill in the Account Number of Beneficiary... ");}
else if(document.f1.agree.checked==0) {document.f1.agree.focus(); alert("Please tick on the Agreement Conditions... ");}
else {document.f1.submit();}
}
