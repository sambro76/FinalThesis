function page(lname,id,wp,pages,count,pn,ordby,mark,hs){
if(pn>pages || pn<1 || isNaN(pn)) {
	alert("Please input an integer number \n or the number should be less than the number of pages");
	}
else {
	url="SBC_PRI.asp?lname=" + lname + "&id=" + id + "&init=1&" + wp + "=1&pages=" + pages + "&count=" + count + "&pageNo=" + pn + "&ordby=" + ordby + "&mark=" + mark + "&hs=" + hs;
	param="toolbar=yes,location=yes,status=yes,scrollbars=yes,menubar=no,resizable=yes,width=800,height=600,left=0,top=0";
	window.open(url,'main', param).opener = self;
}
}
