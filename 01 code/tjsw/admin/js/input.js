function inputStyle(fEvent,oInput){
	if (!oInput.style) return;
	var put=oInput.getAttribute("type").toLowerCase();

	switch (fEvent){
		case "focus" :
			oInput.isfocus = true;
		case "mouseover" :			
			if(put=="submit" || put=="button" || put=="reset")			
				oInput.className="input_on";
			else
				oInput.className = "TextBoxFocus";	
			break;
		case "blur" :
			oInput.isfocus = false;
		case "mouseout" :
			if(put=="submit" || put=="button" || put=="reset")
				oInput.className = "input0";
		    else if(!oInput.isfocus)
				oInput.className = "TextBox";
			break;
		//case else :
			//if(oInput.getAttribute(fEvent+"_2"))
				//eval(oInput.getAttribute(fEvent+"_2"));
	}	
}

window.onload = function(){
	var oInput = document.getElementsByTagName("input");
	var onfocusStr = [];
	var onblurStr = [];
	//alert(oInput.length);
	try
	{
		for (var i=0; i<oInput.length; i++)
		{
			if (!oInput[i]||!oInput[i].getAttribute("type")) continue;
			var put=oInput[i].getAttribute("type").toLowerCase();
			if(put=="submit" || put=="button" || put=="reset")
			{
				oInput[i].className="input0";
			}
			if (put=="text" || put=="password" || put=="submit" || put=="button" || put=="reset")
			{
				if (document.all)
				{
					oInput[i].attachEvent("onmouseover",oInput[i].onmouseover=function(){inputStyle("mouseover",this);});
					oInput[i].attachEvent("onmouseout",oInput[i].onmouseout=function(){inputStyle("mouseout",this);});

				}
				else{
					oInput[i].addEventListener("onmouseover",oInput[i].onmouseover=function(){inputStyle("mouseover",this);},false);
					oInput[i].addEventListener("onmouseout",oInput[i].onmouseout=function(){inputStyle("mouseout",this);},false);				
					//获取焦点
					if(oInput[i].getAttribute("onfocus")){
						oInput[i].addEventListener("onfocus",oInput[i].onblur=function(){eval(this.getAttribute("onfocus"));inputStyle("focus",this);},false);
					}else{
						oInput[i].addEventListener("onfocus",oInput[i].onfocus=function(){inputStyle("focus",this);},false);
					}
					//失去焦点
					if(oInput[i].getAttribute("onblur")){
						oInput[i].addEventListener("onblur",oInput[i].onblur=function(){eval(this.getAttribute("onblur"));inputStyle("blur",this);},false);
					}else{
						oInput[i].addEventListener("onblur",oInput[i].onblur=function(){inputStyle("blur",this);},false);
					}
				}			
			}
		}
	}catch(e){}
	for(i=1;i<=8;i++)//控制面板
	{
		if(document.getElementById('con_two_'+i))
		{	
			document.getElementById('two'+i).className="hover";			
			break;
		}
	}
}