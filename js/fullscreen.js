function MachakFull(Ie,other){
//Copyright ?1999 m.milicevic machakjoe@netscape.net jjooee@tip.nl
x=screen.availWidth;
y=screen.availHeight;
target = parseFloat(navigator.appVersion.substring(navigator.appVersion.indexOf('.')-1,navigator.appVersion.length));
if((navigator.appVersion.indexOf("Mac")!=-1) &&(navigator.userAgent.indexOf("MSIE")!=-1) &&(parseInt(navigator.appVersion)==4))
window.open(other,"sub",'scrollbars=yes');
if (target >= 4){
	if (navigator.appName=="Netscape"){
    var MachakFull=window.open(other,"MachakFull",'scrollbars=yes','width='+x+',height='+y+',top=0,left=0');
	MachakFull.moveTo(0,0);
	MachakFull.resizeTo(x,y);}
if (navigator.appName=="Microsoft Internet Explorer"){
	var MachakFull=window.open(Ie,"MachakFull",'screenX=0,screenY=0,directories=0,width='+x+',height='+y+',location=0,menubar=0,scrollbars=1,status=0,toolbar=0');
	MachakFull.moveTo(0,0);
	MachakFull.resizeTo(x,y);}
	}
	else window.open(other,"sub",'scrollbars=yes');
}
