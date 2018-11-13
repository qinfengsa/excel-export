<!DOCTYPE html> 
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:th="http://www.thymeleaf.org" 
      xmlns:sec="http://www.thymeleaf.org/thymeleaf-extras-springsecurity3"> 
    <head>  
        <meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>${title}</title> 
		<script src="https://code.jquery.com/jquery-3.1.1.min.js"></script>
		<style type="text/css">
			.overLoading {
	            display: none;
	            position: absolute;
	            top: 0;
	            left: 0;
	            width: 100%;
	            height: 100%;
	            background-color: #f5f5f5;
	            opacity:0.5;
	            z-index: 1000;
	        }
		    .layoutLoading {
		        display: none;
		        position: absolute;
		        top: 40%;
		        left: 40%;
		        width: 20%;
		        height: 20%;
		        z-index: 1001;
		        text-align:center;
		    }
		</style> 
    </head> 
    <body> 
    	<div id="over" class="overLoading"></div>
    	<div id="layout" class="layoutLoading"><img src="static/images/loading.gif" /></div>
        <form>
        	<input  type= "button" id = "download" value = "下载Excel"/>
        </form>
        <script> 
	      	//显示loading
	        function showLoading(show){
	        	if (show) {
	        	    $("#over").show();
	        	    $("#layout").show();
	        	} else {
	        		$("#over").hide();
	        	    $("#layout").hide();
	        	}
	        }
        	var w;
			var timer;
			var currentAjax;
			//定时器 当窗口关闭时，中止ajax请求
			function ifWinClosed(){  
			    if(w.closed == true){  
			       	window.clearInterval(timer);
					currentAjax.abort();
			        showLoading(false); 
			    }  
			}
        	function openExportForm(url, params) {  
				var win = window.open("", "_blank");  
				var currentAjax = $.ajax({
					url: url,
					type: "POST",
					data: params,
					dataType: "json",
					cache: false,
					success: function(result) {  
						if (result.success) { 
							//用提交表单的方式下载文件
							var tempForm = $("<form>");
							tempForm.attr("id", "tempForm1");
							tempForm.attr("style", "display:none"); 
							tempForm.attr("method", "post");
							tempForm.attr("action", "/download");  
							var input = $("<input>");       
							input.attr("type", "hidden");
							input.attr("name", "url");
							input.attr("value", result.url);  //下载地址 
							tempForm.append(input);  
							var input2 = $("<input>");       
							input2.attr("type", "hidden");
							input2.attr("name", "excelName");
							input2.attr("value", "MyExcel");  //重命名为MyExcel
							tempForm.append(input2);   
							tempForm.trigger("submit"); 
							$("body").append(tempForm);//将表单放置在web中
							tempForm.submit();
							win.close();
						}  
					} 
				}); 
			
				return {"win":win,"currentAjax":currentAjax}; 
			 
			}
			$("#download").click(function(){
				showLoading(true); 
				var params = {}; 
				var obj = openExportForm("/excel-create",params); 
				w = obj.win;
				currentAjax = obj.currentAjax; 
				timer = window.setInterval('ifWinClosed()',100); 
			})
			
			
        </script>
    </body>

</html>