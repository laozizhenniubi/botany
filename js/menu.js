/*-------------------------------------------\
|      Simple Cross Browser Menu Script      |
|--------------------------------------------|
|        Author:        Emil A. Eklund       |
|        First Created: May 19, 2000         |
|        Last Updated:  Aug 17, 2000         |
|--------------------------------------------|
|     Created to work with ie4+ and ns4+     |
\-------------------------------------------*/

menuPrefix = 'menu';  // Prefix that all menu layers must start with
                      // All layers with this prefix will be treated
                      // as a part of the menu system.

var menuTree, mouseMenu, hideTimer, doHide;
var isIE4=(document.all)?true:false;
var isNS4=(document.layers)?true:false;


function expandMenu(menuContainer,subContainer,menuLeft,menuTop) {
	// 隐藏所有不在该层上的菜单   
	var mouseY,whoami;	
	doHide = false;
  if (menuContainer != menuTree) {
	  if (isIE4) {
      var menuLayers = document.all.tags("DIV");
      for (i=0; i<menuLayers.length; i++) {
        if ((menuLayers[i].id.indexOf(menuContainer) != -1) && (menuLayers[i].id != menuContainer)) {
          unseeObject(menuLayers[i].id);
        }
      }
    }
    else if (isNS4) {
      for (i=0; i<document.layers.length; i++) {
        var menuLayer = document.layers[i];
        if ((menuLayer.id.indexOf(menuContainer) != -1) && (menuLayer.id != menuContainer)) {
          menuLayer.visibility = "hide";
        }
      }
    }
  }
  // If this is item has a submenu, display it, or it it's a toplevel menu, open it
  if (subContainer) {

    if ((menuLeft) && (menuTop)) {
    	positionObject(subContainer,menuLeft,menuTop);
    	hideAll();
    }
    else {
	  whoami=subContainer.substring(subContainer.length-1,subContainer.length)
	  mouseY=123+(whoami-1)*29	
      if (isIE4) {
      	positionObject(subContainer, document.all[menuContainer].offsetWidth + document.all[menuContainer].style.pixelLeft-50, mouseY);
      }
      else {
      	positionObject(subContainer, document.layers[menuContainer].document.width + document.layers[menuContainer].left + 50, mouseY);
      }
    }
    displayObject(subContainer);
    menuTree = subContainer;
  }
}

function displayObject(obj) {
  if (isIE4) { document.all[obj].style.visibility = "visible"; }
  else if (isNS4) { document.layers[obj].visibility = "show";  }
}

function unseeObject(obj) {
  if (isIE4) { document.all[obj].style.visibility = "hidden"; }
  else if (isNS4) { document.layers[obj].visibility = "hide"; }
}

function positionObject(obj,x,y) {
  if (isIE4) {
    var foo = document.all[obj].style;
    foo.left = x;
    foo.top = y;
  }
  else if (isNS4) {
    var foo = document.layers[obj];
    foo.left = x;
    foo.top = y;
   }
}

function hideAll() {
 if (isIE4) {
    var menuLayers = document.all.tags("DIV");
    for (i=0; i<menuLayers.length; i++) {
      if (menuLayers[i].id.indexOf(menuPrefix) != -1) {
        unseeObject(menuLayers[i].id);
      }
    }
  }
  else if (isNS4) {
    for (i=0; i<document.layers.length; i++) {
      var menuLayer = document.layers[i];
      if (menuLayer.id.indexOf(menuPrefix) != -1) {
        unseeObject(menuLayer.id);
      }
    }
  }
}

function hideMe(hide) {
	if (hide) {
		if (doHide) { hideAll(); }
	}
	else {
		doHide = true;
		hideTimer = window.setTimeout("hideMe(true);", 20);
	}
}
