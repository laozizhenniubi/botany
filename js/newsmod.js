<!--
function __off(n)
{
	if(n && n.style)
	{
		if('none' != n.style.display)
		{
			n.style.display = 'none';
		}
	}
}

function __on(n)
{
	if(n && n.style)
	{
		if('none' == n.style.display)
		{
		n.style.display = '';
		}
	}
}


function onoff(objName,bObjState)
{	
	var sVar = ''+objName;
	var sOn  = ''+objName+'_on';
	var sOff = ''+objName+'_off';
	var sOnStyle  =  bObjState ? ' style="display:none;" ':'';
	var sOffStyle = !bObjState ? ' style="display:none;" ':'';
	var sSymStyle = ' style="text-align: center;width: 13;height: 13;font-family: Arial,Verdana;font-size: 7pt;border-style: solid;border-width: 1;cursor: hand;" ';
	if( (navigator.userAgent.indexOf("MSIE") >= 0) && document && document.body && document.body.style)
	{
		document.write( '<span '+sOffStyle+'onclick="__off('+sVar+');__off('+sOff+');__on('+sOn+')" id="'+sOff+'" title="点击这里关闭课程资源模块"'+sSymStyle+'><img src="images/closeleft2.gif" width=13 height=13><\/span>'+
		'<span '+sOnStyle+'onclick="__on('+sVar+');__off('+sOn+');__on('+sOff+')" id="'+sOn+'" title="点击这里打开课程资源模块"'+sSymStyle+'><img src="images/expandleft2.gif" width=13 height=13><\/span>' );
	}
}

// -->
