
//////////////////////////////////////////////////////////////////////
//  Based on overLIB 2.22 from Erik Bosrup (erik@bosrup.com)
//  Last update: 05.04.2001 by Stefan Mezmer (mezmer@betak-design.de)
//////////////////////////////////////////////////////////////////////
var ishover=false;
ns4 = (document.layers);
ie4 = (document.all);
dom = (document.getElementById);
ie6 = false; ns6 = false;
ie5 = (ie4 && navigator.appVersion.indexOf("5.")!=-1);

if (dom) {
  if(navigator.vendor == ("Netscape6") || navigator.product == ("Gecko")) ns6 = true;
  else ie6 = true;
}
if (ie5) { dom = false; ns6 = false; ie6 = false;}

var dir=1;
var x = 0;
var y = 0;
var snow = 0;
var sw = 0;
var cnt = 0;
var tr = 1;
var offsetx = 10;
var offsety = 10;

if ( (ns4) || (ie4) || (dom) ) {
	over = ns4 ? document.overDiv : dom ? document.getElementById("overDiv") : overDiv.style;
	document.onmousemove = mouseMove;
	if (ns4) document.captureEvents(Event.MOUSEMOVE);
}

// Clears popups if appropriate
function nd() {
	ishover=false;
	if ( cnt >= 1 ) { sw = 0 };
	if ( (ns4) || (ie4) || (dom) ) {
		if ( sw == 0 ) {
			snow = 0;
			hideObject(over);
		} else {
			cnt++;
		}
	}
}

// Simple popup
function dts(twidth,backcolor,fcolor,text) {
	txt = "<TABLE WIDTH=\""+twidth+"\" BORDER=0 CELLPADDING=1 CELLSPACING=0 BGCOLOR=\""+backcolor+"\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR=\""+fcolor+"\"><TR><TD><FONT FACE=\"Arial,Helvetica\" COLOR=\"#000000\" SIZE=1>"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>"
	if(text!=0){ishover=true;}
	layerWrite(txt);
	dir = 1;
	disp();
}


// Common calls
function disp() {
	if ( (ns4) || (ie4) || (dom) ) {
		if (snow == 0) 	{
			if (dir == 2) { // Center
				moveTo(over,x+offsetx-(width/2),y+offsety);
			}
			if (dir == 1) { // Right
				moveTo(over,x+offsetx,y+offsety);
			}
			if (dir == 0) { // Left
				moveTo(over,x-offsetx-width,y+offsety);
			}
			showObject(over);
			snow = 1;
		}
	}
// Here you can make the text goto the statusbar.
}

// Moves the layer
function mouseMove(e) {
  if (ishover) {
	if (ns4 || ns6) {x=e.pageX; y=e.pageY;}
	if (ie4) {x=event.x; y=event.y;}
	if (ie5 || ie6) {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
	if (snow) {
		if (dir == 2) { // Center
			moveTo(over,x+offsetx-(width/2),y+offsety);
		}
		if (dir == 1) { // Right
			moveTo(over,x+offsetx,y+offsety);
		}
		if (dir == 0) { // Left
			moveTo(over,x-offsetx-width,y+offsety);
		
		}
	}
  }
}

// The Close onMouseOver function for Sticky
function cClick() {
	hideObject(over);
	sw=0;
}

// Writes to a layer
function layerWrite(txt) {
        if (ns4) {
                var lyr = document.overDiv.document
                lyr.write(txt)
                lyr.close()
        }
        else if (ie4) document.all["overDiv"].innerHTML = txt		
        else if (dom) {
			rng=document.createRange();
			rng.setStartBefore(over);
			htmlFrag = rng.createContextualFragment(txt);
			while (over.hasChildNodes()) { over.removeChild(over.lastChild); }
			over.appendChild(htmlFrag);
	}
}


// Make an object visible
function showObject(obj) {
        if (ns4) obj.visibility = "show";
        else if (ie4) obj.visibility = "visible";
        else if (dom) obj.style.visibility = "visible";
}

// Hides an object
function hideObject(obj) {
        if (ns4) obj.visibility = "hide";
        else if (ie4) obj.visibility = "hidden";
        else if (dom) obj.style.visibility = "hidden";
}

// Move a layer
function moveTo(obj,xL,yL) {
	if ( (ns4) || (ie4) ) {
	        obj.left = xL;
	        obj.top = yL;
	} else if (dom) {
		obj.style.left = xL + "px";
		obj.style.top = yL+ "px";
	}
}