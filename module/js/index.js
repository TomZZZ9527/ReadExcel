let personsSon = [];
let workbookSon = "";

$('#excel-file').change(function(e) {
	var fileInput = $('#excel-file').get(0).files[0];
	var point = (fileInput.name).lastIndexOf(".");
	var type = (fileInput.name).substr(point);
	if(type==".xls"||type==".xlsx"){
	
	$("#ChooseSheet").empty();
	var files = e.target.files;
	var fileReader = new FileReader();
	fileReader.onload = function(ev) {
		try {
			var data = ev.target.result,
				workbook = XLSX.read(data, {
					type: 'binary'
				}), // 以二进制流方式读取得到整份excel表格对象
				persons = []; // 存储获取到的数据
		} catch (e) {
			console.log('文件类型不正确');
			return;
		}
		// 表格的表格范围，可用于判断表头是否数量是否正确
		var fromTo = '';
		var j = "1";
		// 遍历每张表读取
		for (var sheet in workbook.Sheets) {
			if (workbook.Sheets.hasOwnProperty(sheet)) {
				fromTo = workbook.Sheets[sheet]['!ref'];
				//console.log(fromTo);//A1:C4表格范围
				//persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
				//break; // 如果只取第一张表，就取消注释这行
				
				if (j == "1") { //动态创建多选sheet，默认第一个是被选中的，设置为checked="checked"
						$('#ChooseSheet').append('<input lay-filter="sheet" type="checkbox" checked="checked" name="labelType" id="' + "Sheet" + j +
							'"  value="' + "Sheet" + j + '" title="' + "Sheet" + j + '">');
					} else {
						$('#ChooseSheet').append('<input lay-filter="sheet" type="checkbox" name="labelType" id="' + "Sheet" + j + '" value="' + "Sheet" +
							j + '" title="' + "Sheet" + j + '">');
					}
					j = Number(j) + 1;
					form.render("checkbox"); //layer重新渲染
			}
		}
		personsSon = persons;
		workbookSon = workbook;
	}; //文件打开
	// 以二进制方式打开文件
	fileReader.readAsBinaryString(files[0]);
	}else{
		$("#excel-file").val("");
		layer.open({
			title: '提示',
			content: '请重新选择后缀名为xls或xlsx的文件！',
		});
	}
}); //点击方法


function GetData(Sheet) {//根据选择的sheet，然后获取sheet几的数据
	var res = [];
	res = personsSon.concat(XLSX.utils.sheet_to_json(workbookSon.Sheets[Sheet]));
	return res;
}


function RenderExcelSQL() {//渲染穿梭框数据，要是disabled=1就禁选
	for (let i = 0; i < ExcelData.length; i++) {
		if(ExcelData[i].disabled=="1"){
			$('#Excel-Content').append('<input id="Excel'+i+'" lay-filter="Excel" type="radio" name=ExcelRadio value="' + ExcelData[i].title + '" title="' +ExcelData[i].title + 
			'"disabled="disabled">');
		}else{
			$('#Excel-Content').append('<input id="Excel'+i+'" lay-filter="Excel" type="radio" name=ExcelRadio value="' + ExcelData[i].title + '" title="' +ExcelData[i].title + '">');
			}	
	}
	var SQLval="";
	for (let i = 0; i < SQLData.length; i++) {
		SQLval=SQLData[i].excel +SQLData[i].sql 
		if(SQLData[i].disabled=="1"){
			//$('#SQL-Content').append('<input id="SQL'+i+'" lay-filter="SQL" type="radio" name=SQLRadio value="' + SQLval + '" title="' +SQLval + '"disabled="disabled"><button value="' + SQLval + '" class="layui-btn layui-btn-normal" id="Cancel'+i+'" type="button" style="margin-left: -3px;margin-top: 5px;border-radius: 5px;height: 25px;line-height: 25px;font-size: 13px;width: 40px;padding:0px;">取消</button>');
			
		$('#SQL-Content').append('<input id="SQL'+i+'" lay-filter="SQL" type="radio" name=SQLRadio value="' + SQLval + '" title="' +SQLval + '"disabled="disabled"><button value="' + SQLval + '" class="layui-btn layui-btn-normal" id="Cancel'+i+'" type="button" style="background-color:white;color:black;margin-left: -3px;margin-top: 5px;border-radius: 5px;height: 25px;line-height: 25px;font-size: 13px;width: 40px;padding:0px;">取消</button>');
		
		}else{
			$('#SQL-Content').append('<input id="SQL'+i+'" lay-filter="SQL" type="radio" name=SQLRadio value="' + SQLval + '" title="' +SQLval + '">');
			}	
	}
	var form = layui.form;
	form.render("radio");
	for (let i = 0; i < ExcelData.length; i++) {  //判断配对按钮是否高亮
		for (let j = 0; j < SQLData.length; j++) {
			if (ExcelData[i].title == SQLData[j].sql && SQLData[j].excel == "[未设置]") {
				$("#matching").addClass("matching");
			}
		}
	}
	
}
	
	
	
// ------------------------------------------以下是按钮事件
	
	$(document).on('click', '.btn-cursor', function() { //动态生成的class(导入按钮
		cancelExcelVal.push(ExcelVal);
		cancelSQLVal.push(SQLVal);
		var index = SQLVal.lastIndexOf("\]");
		SQLVal = SQLVal.substring(index + 1, SQLVal.length);
		for (let i = 0; i < SQLData.length; i++) {
			if (SQLData[i].sql == SQLVal) { //选择单选框值与单选框的值比较     字段变动[未设置]数据库字段  变为[Excel字段]数据库字段
				SQLData[i].excel = "[" + ExcelVal + "]";
				SQLData[i].disabled = "1"; //导入后禁选
				for (let j = 0; j < ExcelData.length; j++) {
					if (ExcelData[j].title == ExcelVal) {
						ExcelData[j].disabled = "1";
					}
				}
				$("#SQL-Content").empty(); //清除div，重新渲染
				$("#Excel-Content").empty();
				RenderExcelSQL();
			}
		}
		console.log(res);
		SQLchecked = ""; //选择导入后清空点击
		Excelchecked = "";
		$('#btnRight').removeAttr("class", "btn-cursor");
	});
	
	
	
	$("#matching").click(function() { //配对按钮
		for (let i = 0; i < ExcelData.length; i++) {
			for (let j = 0; j < SQLData.length; j++) {
				if (ExcelData[i].title == SQLData[j].sql) {
					//console.log(ExcelData[i].title + "==配对了==" + SQLData[j].sql);
					SQLData[j].excel = "[" + ExcelData[i].title + "]"
					SQLData[j].disabled = "1"
					ExcelData[i].disabled = "1"
				}
			}
		}
		$("#SQL-Content").empty(); //清除div，重新渲染
		$("#Excel-Content").empty();
		RenderExcelSQL();
		$('#matching').removeAttr("class", "matching");
	});
	
	$(document).on('click', '.layui-btn-normal', function(e) { //取消按钮
		var btnVal = $(e.target).attr('value');
		var excelVal = btnVal.substring(btnVal.indexOf("[") + 1, btnVal.indexOf("]"))
		var index = btnVal.lastIndexOf("\]");
		SQLVal = btnVal.substring(index + 1, btnVal.length);
		for (let i = 0; i < ExcelData.length; i++) {
			if (ExcelData[i].title == excelVal) {
				ExcelData[i].disabled = "0";
			}
		}
		for (let i = 0; i < SQLData.length; i++) {
			if (SQLData[i].sql == SQLVal) {
				SQLData[i].excel = "[未设置]";
				SQLData[i].disabled = "0";
			}
		}
		$("#SQL-Content").empty(); //清除div，重新渲染
		$("#Excel-Content").empty();
		RenderExcelSQL();
	});