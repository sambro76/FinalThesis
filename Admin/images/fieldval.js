function validate(){
	if(document.f1.AccID.value.length<6){window.alert("The login account should not be less 6 characters! "); document.f1.AccID.select();return 0;}
	else if(document.f1.pwd.value.length<3){window.alert("The password size is small. It must be at least 3 characters... "); document.f1.pwd.select(); return 0;}
	else if(document.f1.pwd.value!=document.f1.conpwd.value){window.alert("Password verification failed in Confirm Password... "); document.f1.conpwd.select(); return 0;}
	else if(document.f1.InitCredit.value<500){window.alert("The Initial Credit Amount should be at least 500$... "); document.f1.InitCredit.select(); return 0;}
	else if(document.f1.FName.value.length<2){window.alert("Please verify the First Name Field... "); document.f1.FName.select(); return 0;}
	else if(document.f1.LName.value.length<2){window.alert("Please verify the Last Name Field... "); document.f1.LName.select(); return 0;}
	else if(document.f1.AccNo.value.length<6){window.alert("Please input the certified bank account number... "); document.f1.AccNo.select(); return 0;}
	else if(document.f1.HB.value==""){window.alert("Please fill in House/Building No.... "); document.f1.HB.select(); return 0;}
	else if(document.f1.Street.value==""){window.alert("Please fill in Street field... "); document.f1.Street.select(); return 0;}
	else if(document.f1.City.value.length<3){window.alert("Please fill in the correct City/State... "); document.f1.City.select(); return 0;}
	else if(document.f1.Country.value.length<3){window.alert("Please fill in the correct Country... "); document.f1.Country.select(); return 0;}
	else {return 1;}
}