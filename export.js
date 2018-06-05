/*
 * author:xyz
 */

var xyz_export = function () {
    let styles;
    let def_opt = {
        rowHeight: 13.5,
        headHeight: 0,
        fontColor:"#000000",
        border: {
            style: "",
            color: ""
        }
    };
    let border_style={
        "default": "Continuous,0",
        "default1": "Continuous,1",
        "default2": "Continuous,2",
        "default3": "Continuous,3",
        "dot":"Dot,1",
        "dash": "Dash,1",
        "dash1": "Dash,2",
        "dashDot": "DashDot,1",
        "dashDot1": "DashDot,2",
        "dashDots": "DashDotDot,1",
        "dashDots1": "DashDotDot,2",
        "slantDashDot":"SlantDashDot,2",
        "double":"Double,3"
    }

    function createEle(type) {
        return document.createElementNS("urn:schemas-microsoft-com:office:spreadsheet", type);
    }

    function _opt(s, o) {
        let tmp = {};
        for (let name in o)
            tmp[name] = o[name];
        for (let name in s) {
            if (typeof (tmp[name]) == "object" && typeof (s[name]) == "object")
                tmp[name]=_opt(s[name], tmp[name]);
            else
                if (s[name] == tmp[name]) continue;
                else
                    tmp[name] = s[name];
        }
        return tmp;
    }

    function sheet(obj) {
        styles = [];
        if (obj.data.length < 1 || obj.sheets.length < 1)
            throw "setting error,Please ensure that the length of data and sheet is greater than 1";
        if (obj.data.length < obj.sheets.length)
            throw "The data length must be less than or equal to the length of the worksheet";
        let sheet_list = generate_sheet(obj.data, obj.sheets);
        file = generate_file(sheet_list);
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
         <Styles>
         ${styles.join("")}
         </Styles>
        ${bd}
        </Workbook>`;
        return template;
    }

    function generate_style(s, k) {
        let border;
        let fontColor = s.fontColor;
        if (border_style[s.border.style]){
            let style = border_style[s.border.style].split(",")[0];
            let width = border_style[s.border.style].split(",")[1];
            let color = s.border.color || "#000000";
            border = `<Border ss:Position="Bottom" ss:LineStyle="${style}" ss:Weight="${width}" ss:Color="${color}"/>
                <Border ss:Position="Left" ss:LineStyle="${style}" ss:Weight="${width}"  ss:Color="${color}"/>
                <Border ss:Position="Right" ss:LineStyle="${style}" ss:Weight="${width}"  ss:Color="${color}"/>
                <Border ss:Position="Top" ss:LineStyle="${style}" ss:Weight="${width}"  ss:Color="${color}"/>`;
        }
        else
            boder = "";
        let style = `<Style ss:ID="cs${k}">
            <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
            <Borders>
            ${border}
            </Borders>
            <Font ss:FontName="宋体" x:CharSet="134" ss:Size="12" ss:Color="${fontColor}"/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
            </Style>`;
        styles.push(style);
        debugger;
    }

    

    function generate_sheet(data,sheets) {
        let T_sheet = createEle("Worksheet");
        let T_table = createEle("Table");
        let T_column = createEle("Column");
        let T_row = createEle("Row");
        let T_cell = createEle("Cell");
        let T_D = createEle("Data");
        T_D.setAttribute("ss:Type", "String");
        let th = T_row.cloneNode();
        
        let sheet_list = [];
        let sheets_length=sheets.length;
        for (let k = 0; k < sheets_length; k++) {
            let opt;
            let sheet = T_sheet.cloneNode();
            let table = T_table.cloneNode();
            if (sheets[k].style) {
                opt = _opt(sheets[k].style, def_opt)
                if (opt.border) generate_style(opt, k);
            };

            sheet.setAttribute("ss:Name", sheets[k].sheetName ? sheets[k].sheetName : sheets[k].title ? sheets[k].title.text ? sheets[k].title.text : ("sheet" + (k + 1)) : ("sheet" + (k + 1)));
            table.setAttribute("ss:ExpandedColumnCount", sheets[k].body.length);
            table.setAttribute("ss:ExpandedRowCount", data[k].length + 1 + (sheets[k].title ? 1 : 0));
            table.setAttribute("ss:DefaultColumnWidth", "54");
            table.setAttribute("ss:DefaultRowHeight", opt.rowHeight);
            
            let body = sheets[k].body;
            let body_length = body.length;
            let data_length = data[k].length;
            
            for (let i = 0; i < data[k].length; i++) {
                let row = T_row.cloneNode();
                for (let j = 0; j < body_length; j++) {
                    if (i == 0) {
                        let cell = T_cell.cloneNode();
                        cell.setAttribute("ss:StyleID", "cs"+k);
                        cell.appendChild(T_D.cloneNode()).innerHTML = body[j].text?body[j].text:("column"+(j+1));
                        th.appendChild(cell);//create table head
                        if (body[j].width) {
                            let column = T_column.cloneNode();
                            if (j == 0)
                                column.setAttribute("ss:Width", body[j].width);
                            else
                                if (body[j - 1].width)
                                    column.setAttribute("ss:Width", body[j].width);
                                else {
                                    column.setAttribute("ss:Index", j + 1);
                                    column.setAttribute("ss:Width", body[j].width);
                                }
                            table.appendChild(column);//append column set
                        }
                    }
                    let cell = T_cell.cloneNode();
                    cell.setAttribute("ss:StyleID", "cs"+k);
                    cell.appendChild(T_D.cloneNode()).innerHTML = data[k][i][body[j].field ? body[j].field : function () { throw "Field is null";}()]?data[k][i][body[j].field]:"";
                    row.appendChild(cell);//create data row
                }
                if (i == 0) {
                    if (sheets[k].title) {
                        let row = T_row.cloneNode();
                        let cell = T_cell.cloneNode();

                        row.setAttribute("ss:Height", sheets[k].title.height ? sheets[k].title.height : 50);
                        cell.setAttribute("ss:StyleID", "cs"+k);
                        cell.setAttribute("ss:MergeAcross", sheets[k].body.length - 1);
                        cell.appendChild(T_D.cloneNode()).innerHTML = sheets[k].title.text ? sheets[k].title.text : sheets[k].sheetName?sheets[k].sheetName:("sheet" + (k + 1));
                        row.appendChild(cell);
                        table.appendChild(row);//append  table title
                    }
                    if (opt.headHeight != 0)
                        th.setAttribute("ss:Height", opt.headHeight);
                    table.appendChild(th);//append table head
                }
                table.appendChild(row);//append data row on the table
            }
            sheet.appendChild(table);
            sheet_list.push(sheet.outerHTML);
        }
        return sheet_list;
    }

    return {
        sheet:sheet
    }
}();
