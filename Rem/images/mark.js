<!--
function anySelected(row){
	for(var i=1; i<=row; i++){
		if(document.f1.elements[i+4].checked) {return true;}
	}
	return false;
}
function mark(row, lname, id, sel, pages, count, pageNo, ordby, hs){
	if (anySelected(row)) {
		var url="mark.asp?" + "lname=" + lname + "&id=" + id + "&row=" + row;
		for(i=1; i<=row; i++){
			if(document.f1.elements[i+4].checked){
				url = url + "&cb[" + i + "]=" + document.f1.elements[i+4].value;
			}
		}
		if(sel=="r"){url=url + "&sel=1";}
		else if(sel=="u"){url=url + "&sel=0";}
		else if(sel=="d1") {url=url + "&sel=2";}
		else if(sel=="d0") {url=url + "&sel=1";}
		url=url + "&pages=" + pages + "&count=" + count + "&pageNo=" + pageNo + "&ordby=" + ordby + "&hs=" + hs;
		window.open(url,'main');
	}
	else {alert("You must select at least one form first... "); document.f1.reset();}
}
//-->