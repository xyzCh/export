/*
 * author:xyz
 * date:2018-11-21
 */

var Exl = function () {

	let def_BodyOpt = {
		headHeight: 15,
		rowHeight: 13.5,
		colWidth: 54,
		fontColor: "#000000",
		fontSize: 12,
		wrapText: true,
		border: {
			style: "",
			color: ""
		}
	};
	let def_TitleOpt = {
		text: "",
		height: 50,
		fontColor: "#000000",
		fontSize: 20,
		wrapText: true,
		border: {
			style: "",
			color: ""
		}
	};
	let border_styles = {
		"default": "Continuous,0",
		"default1": "Continuous,1",
		"default2": "Continuous,2",
		"default3": "Continuous,3",
		"dot": "Dot,1",
		"dash": "Dash,1",
		"dash1": "Dash,2",
		"dashDot": "DashDot,1",
		"dashDot1": "DashDot,2",
		"dashDots": "DashDotDot,1",
		"dashDots1": "DashDotDot,2",
		"slantDashDot": "SlantDashDot,2",
		"double": "Double,3"
	};

	function createEle(type) {
		return document.createElementNS("urn:schemas-microsoft-com:office:spreadsheet", type);
	}

	function _opt(s, o) {
		let tmp = {};
		for (let name in o)
			tmp[name] = o[name];
		for (let name in s) {
			if (typeof (tmp[name]) == "object" && typeof (s[name]) == "object")
				tmp[name] = _opt(s[name], tmp[name]);
			else
				if (s[name] == tmp[name]) continue;
				else
					tmp[name] = s[name] || tmp[name];
		}
		return tmp;
	}

	function copy(arr) {
		let res = [];
		for (let i = 0; i < arr.length; i++) {
			if (arr[i] instanceof Array)
				res[i] = copy(arr[i]);
			else
				res[i] = arr[i];
		}
		return res;
	}

	function sheet(obj) {
		if (obj.data.length < 1 || obj.sheets.length < 1)
			throw "setting error,Please ensure that the length of data and sheet is greater than 1";
		if (obj.data.length < obj.sheets.length)
			throw "The data length must be less than or equal to the length of the worksheet";
		let sheet_list = generate_sheet(obj.data, obj.sheets);
		let file = generate_file(sheet_list);
		let a = document.createElement('a');
		a.href = 'data:application/vnd.ms-excel;base64,' + window.btoa(unescape(encodeURIComponent(file)));
		a.download = obj.fileName ? obj.fileName : "exportSheet";
		a.click();
	}

	function generate_file(bd) {
		let template = `<?xml version="1.0"?>
        <?mso-application progid="Excel.Sheet"?>
        <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:html="http://www.w3.org/TR/REC-html40">
        ${bd}
        </Workbook>`;
		return template;
	}



	let T_styles = createEle("Styles");
	let T_style = createEle("Style");
	let T_borders = createEle("Borders");
	let T_border = createEle("Border");
	let T_sheet = createEle("Worksheet");
	let T_table = createEle("Table");
	let T_column = createEle("Column");
	let T_row = createEle("Row");
	let T_cell = createEle("Cell");
	let T_th = T_row.cloneNode();
	let T_D = createEle("Data");
	T_D.setAttribute("ss:Type", "String");


	function generate_style(s, k, styles, kind) {
		let border;
		let fontColor = s.fontColor;
		let fontSize = s.fontSize;
		let WrapText = s.wrapText ? 1 : 0;
		let style = T_style.cloneNode();
		let alignment = createEle("Alignment");
		alignment.setAttribute("ss:Horizontal", "Center");
		alignment.setAttribute("ss:Vertical", "Center");
		alignment.setAttribute("ss:WrapText", WrapText);
		style.appendChild(alignment);
		if (border_styles[s.border.style]) {
			let border_style = border_styles[s.border.style].split(",")[0];
			let border_width = border_styles[s.border.style].split(",")[1];
			let border_color = s.border.color || "#000000";
			let position = ["Bottom", "Left", "Right", "Top"];
			let borders = T_borders.cloneNode();
			for (let i = 0; i < 4; i++) {
				let border = T_border.cloneNode();
				border.setAttribute("ss:Position", position[i]);
				border.setAttribute("ss:LineStyle", border_style);
				border.setAttribute("ss:Weight", border_width);
				border.setAttribute("ss:Color", border_color);
				borders.appendChild(border);
				style.appendChild(borders);
			}
		}
		let font = createEle("Font");
		font.setAttribute("ss:FontName", "宋体");
		font.setAttribute("x:CharSet", "134");
		font.setAttribute("ss:Size", fontSize);
		font.setAttribute("ss:Color", fontColor);
		style.appendChild(font);
		style.setAttribute("ss:ID", "cs" + k + (kind ? "_" + kind : ""));
		styles.appendChild(style);
	}

	function generate_sheet(data, sheets) {
		let sheet_list = [];
		let styles = T_styles.cloneNode();
		let sheets_length = sheets.length;
		for (let k = 0; k < sheets_length; k++) {
			let sheet = T_sheet.cloneNode();
			let table = T_table.cloneNode();
			let _body = copy(sheets[k].body);
			if (sheets[k].rowNumber)
				_body[0].unshift({ field: "rownumber", text: "序号", width: "30", rowspan: sheets[k].body.length });
			let s_body = spreadBody(copy(_body));
			let _bodyopt = _opt(sheets[k].style || {}, def_BodyOpt);
			let _titleopt = _opt(sheets[k].title || {}, def_TitleOpt);
			generate_style(_bodyopt, k, styles);

			generate_Column(s_body, table);
			if (sheets[k].title) {
				generate_style(_titleopt, k, styles, "title");
				generate_Title(s_body.length, _titleopt, sheets[k].sheetName, k, table);
			}
			if (sheets[k].showHeadRow)
				generate_Head(copy(_body), _bodyopt, k, table);
			let datarows = generate_Data(s_body, data[k], table, k, sheets[k].rowNumber);
			sheet.setAttribute("ss:Name", sheets[k].sheetName ? sheets[k].sheetName : _titleopt.text ? _titleopt.text : ("sheet" + (k + 1)));
			table.setAttribute("ss:ExpandedColumnCount", s_body.length);
			table.setAttribute("ss:ExpandedRowCount", datarows + (sheets[k].showHeadRow ? sheets[k].body.length : 0) + (sheets[k].title ? 1 : 0));
			table.setAttribute("ss:DefaultColumnWidth", _bodyopt.colWidth);
			table.setAttribute("ss:DefaultRowHeight", _bodyopt.rowHeight);
			sheet.appendChild(table);
			sheet_list.push(sheet.outerHTML);
		}
		sheet_list.unshift(styles.outerHTML);
		return sheet_list;
	}

	function spreadBody(body) {
		return pool(body);
		function pool(bd) {
			let rowIndex = arguments[1] ? arguments[1] : 0;
			let readWidth = arguments[2] ? arguments[2] : 0;
			let creadWidth = 0;
			let cols = [];
			let row = bd[rowIndex];
			let requiredSpan = bd.length - rowIndex;
			for (let i = 0; i < row.length;) {
				let rowspan = ~~row[i].rowspan == 0 ? 1 : ~~row[i].rowspan;
				let colspan = ~~row[i].colspan == 0 ? 1 : ~~row[i].colspan;
				if (rowspan < requiredSpan) {
					cols = cols.concat(pool(bd, rowIndex + rowspan, colspan));
				} else {
					cols.push(row[i]);
				}
				row.shift();
				creadWidth += colspan;
				if (creadWidth == readWidth)
					break;
			}
			return cols;
		}
	}

	function generate_Column(body, table) {
		for (let i = 0; i < body.length; i++) {
			let column = T_column.cloneNode();
			if (body[i].width)
				column.setAttribute("ss:Width", body[i].width);
			table.appendChild(column);
		}
	}

	function generate_Title(merge, opt, sheetName, k, table) {
		let row = T_row.cloneNode();
		let cell = T_cell.cloneNode();
		let data = T_D.cloneNode();
		data.innerHTML = opt.text ? opt.text : sheetName ? sheetName : ("sheet" + (k + 1));
		cell.appendChild(data);
		cell.setAttribute("ss:StyleID", "cs" + k + "_title");
		cell.setAttribute("ss:MergeAcross", merge - 1);
		row.appendChild(cell);
		row.setAttribute("ss:Height", opt.height);
		table.appendChild(row);
	}

	function generate_Head(body, opt, k, table) {
		let cols = [], cursor = 1;
		for (let i = 0; i < body.length; i++) {
			let row = T_row.cloneNode();
			row.setAttribute("ss:Height", opt.headHeight);
			table.appendChild(row);
			cols[i] = row;
		}
		pool(body);
		return cols;
		function pool(bd) {
			let rowIndex = arguments[1] ? arguments[1] : 0;
			let readWidth = arguments[2] ? arguments[2] : 0;
			let creadWidth = 0;
			let row = bd[rowIndex];
			let requiredSpan = bd.length - rowIndex;
			for (let i = 0; i < row.length;) {
				let cell = T_cell.cloneNode();
				let data = T_D.cloneNode();
				let rowspan = ~~row[i].rowspan == 0 ? 1 : ~~row[i].rowspan;
				let colspan = ~~row[i].colspan == 0 ? 1 : ~~row[i].colspan;
				data.innerHTML = row[i].text || "column" + cursor;
				cell.appendChild(data);
				cell.setAttribute("ss:StyleID", "cs" + k);
				if (rowspan > 1) {
					cell.setAttribute("ss:MergeDown", rowspan - 1);
				}
				if (colspan > 1) {
					cell.setAttribute("ss:MergeAcross", colspan - 1);
				}
				if (creadWidth == 0)
					cell.setAttribute("ss:Index", cursor);
				cols[rowIndex].appendChild(cell);
				row.shift();
				if (rowspan < requiredSpan) {
					pool(bd, rowIndex + rowspan, colspan);
				} else {
					cursor++;
				}
				creadWidth += colspan;
				if (creadWidth == readWidth)
					break;
			}
		}
	}

	function generate_Data(body, data, table, k, rownum) {
		let pre_val = [], mergeRow_list = [], mergeCell_list = [];
		let rows = data.length;
		for (let z = 0; z < body.length; z++) {
			pre_val.push("");
			mergeRow_list.push(0);
		}

		for (let i = 0; i < data.length; i++) {
			let row = T_row.cloneNode();
			let j = 0;
			if (rownum) {
				j = 1;
				let cell = T_cell.cloneNode();
				cell.setAttribute("ss:StyleID", "cs" + k);
				cell.appendChild(T_D.cloneNode()).innerHTML = (i + 1);
				row.appendChild(cell);
			}
			for (; j < body.length; j++) {
				let cell = T_cell.cloneNode();
				cell.setAttribute("ss:StyleID", "cs" + k);
				if (j > 0 && body[j - 1].merge)
					cell.setAttribute("ss:Index", j + 1);
				let val = data[i][body[j].field ? body[j].field : function () { throw "Field is null"; }()] ? data[i][body[j].field] : "";
				if (body[j].merge) {
					if (pre_val[j] != val) {
						pre_val[j] = val;
						if (mergeRow_list[j] != 0) {
							mergeCell_list[j].setAttribute("ss:MergeDown", mergeRow_list[j]);
						}
						mergeCell_list[j] = cell;
						mergeRow_list[j] = 0;
					} else {
						mergeRow_list[j] += 1;
						if (i == data.length - 1 && mergeRow_list[j] != 0)
							mergeCell_list[j].setAttribute("ss:MergeDown", mergeRow_list[j]);
						continue;
					}
				}
				cell.appendChild(T_D.cloneNode()).innerHTML = body[j].formatter ? body[j].formatter(val, i, data[i]) : val;
				row.appendChild(cell);
			}
			table.appendChild(row);
		}
		return rows;
	}

	return {
		sheet: sheet
	}
}();
