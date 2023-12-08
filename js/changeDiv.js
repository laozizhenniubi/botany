function DisplayObject(obj) {
  if (isIE4) { document.all[obj].style.display = ''; }
  else if (isNS4) { document.layers[obj].display  = '';  }
}
function NoneObject(obj) {
  if (isIE4) { document.all[obj].style.display= 'none'; }
  else if (isNS4) { document.layers[obj].display  = 'none';  }
}
function MM_openBrWindow(theURL,winName,features) { 
  window.open(theURL,winName,features);
}