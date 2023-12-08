//var isIE4=(document.all)?true:false;
//var isNS4=(document.layers)?true:false;
var display_url=0;

var menuobj=document.getElementById("ie5menu");

function showmenuie5(e){

	var rightedge=isIE4? document.body.clientWidth-event.clientX : window.innerWidth-e.clientX;
	var bottomedge=isIE4? document.body.clientHeight-event.clientY : window.innerHeight-e.clientY;


	if (rightedge<menuobj.offsetWidth) { menuobj.style.left=isIE4? document.body.scrollLeft+event.clientX-menuobj.offsetWidth : window.pageXOffset+e.clientX-menuobj.offsetWidth;
	}
	else{
	menuobj.style.left=isIE4? document.body.scrollLeft+event.clientX : window.pageXOffset+e.clientX;
	}


if (bottomedge>menuobj.offsetHeight) { menuobj.style.top=isIE4? document.body.scrollTop+event.clientY-menuobj.offsetHeight : window.pageYOffset+e.clientY-menuobj.offsetHeight;
}
else{
menuobj.style.top=isIE4? document.body.scrollTop+event.clientY : window.pageYOffset+e.clientY;
}

menuobj.style.visibility="visible";
return false;
}

function hidemenuie5(e){
menuobj.style.visibility="hidden";
}

function highlightie5(e){
var firingobj=isIE4? event.srcElement : e.target;
if (firingobj.className=="menuitems"||isNS4&&firingobj.parentNode.className=="menuitems"){
	if (isNS4&&firingobj.parentNode.className=="menuitems") firingobj=firingobj.parentNode; //up one node
	firingobj.style.backgroundColor="highlight";
	firingobj.style.color="white";
	if (display_url==1) window.status=event.srcElement.url;
	}
}

function lowlightie5(e){
	var firingobj=isIE4? event.srcElement : e.target;
	if (firingobj.className=="menuitems"||isNS4&&firingobj.parentNode.className=="menuitems"){
	if (isNS4&&firingobj.parentNode.className=="menuitems") firingobj=firingobj.parentNode; //up one node
	firingobj.style.backgroundColor="";
	firingobj.style.color="black";
	window.status='';
	}
}

function jumptoie5(e){
	var firingobj=isIE4? event.srcElement : e.target;
	if (firingobj.className=="menuitems"||isNS4&&firingobj.parentNode.className=="menuitems"){
	if (isNS4&&firingobj.parentNode.className=="menuitems") firingobj=firingobj.parentNode;
	if (firingobj.getAttribute("target"))
	window.open(firingobj.getAttribute("url"),firingobj.getAttribute("target"));
	else
	window.location=firingobj.getAttribute("url");
	}
}

if (isIE4||isNS4){
menuobj.style.display='';
document.oncontextmenu=showmenuie5;
document.onclick=hidemenuie5;
}
