listOutSpeed=16;
function gotoUrl(sUrl){
	window.open(sUrl,"","");
}
function listOut(){
	if(selectedCon.listState=="close") {
		selectedCon.listState="opening";
		lListMask.style.visibility="visible";
		lListMask.setCapture();
		doListOut();
	}
}
function doListOut(){
	lList.style.pixelTop+=listOutSpeed;
	if (lList.style.pixelTop<0) setTimeout("doListOut()",1)
		else {
			lList.style.pixelTop=0;
			selectedCon.listState="open";
			}
}
function doSelect(){
var o=event.srcElement;
	lListMask.style.visibility="hidden";
	lList.style.pixelTop-=lList.style.pixelHeight;
	selectedCon.listState="close";
	lListMask.releaseCapture();
	if(o.className=="listText2") {
	selectedIcon.src=o.parentElement.children(0).children(0).src;
	selectedTxt.innerText=o.innerText;
	gotoUrl(o.value);
	}
}

function doMouseOver(){
var o=event.srcElement;
	if(o.className=="listText") o.className="listText2";
}

function doMouseOut(){
var o=event.srcElement;
	if(o.className=="listText2") o.className="listText";
}

document.onmouseover=doMouseOver;
document.onmouseout=doMouseOut;