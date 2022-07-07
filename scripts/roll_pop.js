<!--
function popwindow(url,w,h) { //klee v1.0
var sw=screen.availWidth;sh=screen.availHeight;l=(sw-w)/2;t=(sh-h)/2;
window.open(url,'lcpop','width='+w+',height='+h+',left='+l+',top='+t+',screenX='+l+',screenY='+t+',resizable,scrollbars');
}

function PRO_openPopupWindow(theURL,winName,intWidth,intHeight,features,centralise) { //pro v2.0
	features = features + ",width=" + intWidth + ",height=" + intHeight;

	if (centralise == "yes") {
		var intAvailWidth = 640, intAvailHeight = 480;
		var intMargin = 10;
		var intTop = intMargin, intLeft = intMargin;
		if (typeof(screen) == "object") {
			intAvailWidth = screen.availWidth;
			intAvailHeight = screen.availHeight;
		}
		intTop = Math.round(intAvailHeight/2 - intHeight/2);
		if (intTop < intMargin) intTop = intMargin;
		intLeft = Math.round(intAvailWidth/2 - intWidth/2);
		if (intLeft < intMargin) intLeft = intMargin;
		features = features + ",left=" + intLeft + ",top=" + intTop;
	}

  var newWin = window.open(theURL,winName,features);
  if (newWin.focus) newWin.focus();
  document.MM_returnValue = false;
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->