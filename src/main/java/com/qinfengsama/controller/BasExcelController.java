package com.qinfengsama.controller;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import com.qinfengsama.utils.ExcelUtils;

@Controller
public class BasExcelController {
	
	private Logger logger = LoggerFactory.getLogger(this.getClass());
	
	/**
	 * excel界面
	 * @param map
	 * @return
	 */
	@RequestMapping("/helloExcel") 
    public String getExcelHtml(Map<String,Object> map){

		map.put("title","Excel导出测试");

       	return"/helloExcel";

    }
	@RequestMapping("/") 
    public String index(Map<String,Object> map){
 
		return "redirect:/helloExcel";

    }
	/**
	 ** 创建excel文件
	 * @param request
	 * @param response
	 * @return
	 * @throws InterruptedException 
	 */
	@RequestMapping(value = "/excel-create", method = RequestMethod.POST)
	@ResponseBody
	public Map<String,Object> exportSporder(HttpServletRequest request,HttpServletResponse response) throws InterruptedException {
		Map<String, Object> resultMap = new HashMap<String, Object>();
	 
		List<Map<String, Object>> lists = new ArrayList<Map<String,Object>>();
		for (int i = 0; i < 5; i++) {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("name1" , UUID.randomUUID().toString());
			map.put("name2" , UUID.randomUUID().toString());
			map.put("name3" , UUID.randomUUID().toString());
			map.put("name4" , UUID.randomUUID().toString());
			map.put("name5" , UUID.randomUUID().toString());
			map.put("name6" , UUID.randomUUID().toString());
			map.put("name7" , UUID.randomUUID().toString());
			map.put("name8" , UUID.randomUUID().toString());
			map.put("name9" , UUID.randomUUID().toString());
			map.put("name10" , UUID.randomUUID().toString()); 
			
			lists.add(map);
			
		}
		String[] tableHeader = {"名称1", "名称2", "名称3","名称4","名称5",
				"名称6","名称7","名称8","名称9","名称10" }; 
		String[] tableProperties = {"name1", "name2", "name3","name4","name5",
				"name6","name7","name8","name9","name10" };
		//5秒后执行，验证loading和中断效果
		Thread.currentThread().sleep(5000);
		try {
			String url = ExcelUtils.createExcelSheetUrl(lists, tableHeader, tableProperties,"test",request);
		 
			resultMap.put("success", true);
			resultMap.put("errorMsg","");
			resultMap.put("url", url);
		} catch (Exception e) {
			resultMap.put("success", false);
			resultMap.put("errorMsg", e.getMessage());
			resultMap.put("url", "");
			logger.error(e.getMessage(),e);
		} 
		return resultMap;
	}
	 
	/**
	 * 
	 * @Title: excelDownLoad  
	 * @Description: TODO(把生成的文件下载并重命名为excelName)  
	 * @param url
	 * @param excelName
	 * @param request
	 * @param response 
	 * @throws IOException     
	 * @return ModelAndView    返回类型
	 */
	@RequestMapping(value = "/download", method = RequestMethod.POST)
	public ModelAndView excelDownLoad(String url,String excelName,HttpServletRequest request,HttpServletResponse response) throws IOException {
	 
		String realPath = request.getSession().getServletContext().getRealPath("/") + url;
        File file = new File(realPath);
		if(!file.exists()) {
			throw new RuntimeException("文件不存在!");
		}
		FileInputStream fileInputStream = new FileInputStream(file);
		BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
		OutputStream outputStream = response.getOutputStream();
		BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(outputStream);
 
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/x-download"); 
		response.setHeader("Content-disposition", "attachment;filename=" + new String((excelName + ".xlsx").getBytes("UTF-8"), "iso-8859-1"));
		int bytesRead = 0;
		byte[] buffer = new byte[8192];
		while ((bytesRead = bufferedInputStream.read(buffer, 0, 8192)) != -1) {
			bufferedOutputStream.write(buffer, 0, bytesRead);
		}
		bufferedOutputStream.flush();
		fileInputStream.close();
		bufferedInputStream.close();
		outputStream.close();
		bufferedOutputStream.close();
        if(file.exists()){
            file.delete();//删除excel文件,反正之后没用了
        }
		return null;
	}

}
