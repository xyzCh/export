export
===
`导出excel表格...很简陋 :-) `

Docs
---
**用法**  
```javascript
Exl.sheet({
    fileName:"file",
    data:[[{a:1,b:2,c:3,d:4,e:5,f:6},{a:2,b:2,c:5,d:4,e:5,f:6}...],...],
    sheets:[
        {
            sheetName:"sheet1",
            rowNumber:true,
	    showHeadRow:true,
            title: {
                text: "test",
                height: 50,
		fontColor:"#000000",
		fontSize:20,
		wrapText:true,
		border: {
		    style: "",
		    color: ""
		}
            },
            style:{
                headHeight:15,
                rowHeight:13.5,
	     	colWidth:54,
                fontColor:"#000000",
		fontSize:12,
		wrapText:true,
                border:{
                    style:"default1",
                    color:"#000000"
                }
            },
            body: [[
                { field: "a", text: "字段1", width: 50, rowspan:3, formatter:function(value,index,rowData){
                        .....
                        return value;
                    }
                },
                { field: "b", text: "字段2", width: 50, rowspan:3, merge:true },//此列上相同的数据(相邻行)会发生合并
                { field: "c", text: "字段3", width: 50, rowspan:3 },
                { text:"合并", colspan:3 },
                { text:"合并2", colspan:2, rowspan:2 }
                ],[
                    { field:"d", text:"字段4", width: 50, rowspan:2 },
                    { text:"合并1", colspan:2 }
                ],[
                    { field:"e", text:"字段5", width: 50, rowspan:2 },
                    { field:"f", text:"字段6", width: 50, rowspan:2 },
                    { field:"g", text:"字段7", width: 50, rowspan:2 },
                    { field:"h", text:"字段8", width: 50, rowspan:2 }
                ]
            ]
        },...
    ]
});
```
**配置**

|名称|类型|描述|默认值|
|:--:|:--:|:--:|:--:|
|fileName|string|导出的文件名|exportSheet.xls|
|data|array|要导出的数据|undefined|
|[sheets](#sheet)|array|[sheet](#sheet)设置的集合|undefined|

#### sheet
<table>
    <tr>
        <th>名称</th>
        <th>类型</th>
        <th>描述</th>
        <th></th>
    </tr>
    <tr>
        <td>sheetName</td>
        <td>string</td>
        <td>每个工作表的名字</td>
        <td>如果sheetName未设置,但title.text已设置,则默认title.text,当两者都未设置则为sheet[1,2,...]</td>
    </tr>
    <tr>
        <td>rowNumber</td>
        <td>boolean</td>
        <td>是否显示行号</td>
        <td></td>
    </tr>
	<tr>
        <td>showHeadRow</td>
        <td>boolean</td>
        <td>是否显示表头</td>
        <td></td>
    </tr>
    <tr>
        <td>title</td>
        <td>object</td>
        <td>工作表标题设置</td>
        <td>
            不设置title时,默认不显示sheet标题<br/>
            <table>
                <tr>
                    <td>title.text</td>
                    <td>标题名字</td>
                    <td>如果text未设置,但sheetName已设置,则text默认为sheetName,当两者都未设置则为sheet[1,2,...]</td>
                </tr>
                <tr>
                    <td>title.height</td>
                    <td>标题高度</td>
                    <td>默认50</td>
                </tr>
                <tr>
                    <td>title.fontColor</td>
                    <td>标题字体颜色</td>
                    <td>默认#000000</td>
                </tr>
                <tr>
                    <td>title.fontSize</td>
                    <td>标题字体大小</td>
                    <td>默认20</td>
                </tr>
                <tr>
                    <td>title.wrapText</td>
                    <td>自动换行</td>
                    <td>默认true</td>
                </tr>
                <tr>
                    <td>title.border</td>
                    <td>标题边框</td>
                    <td>同style.border</td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>style</td>
        <td>object</td>
        <td>工作表样式设置</td>
        <td>
            <table>
                <tr>
                    <td>style.headHeight</td>
                    <td>表头高度</td>
                    <td>默认15</td>
                </tr>
                <tr>
                    <td>style.rowHeight</td>
                    <td>行高</td>
                    <td>默认13.5</td>
                </tr>
                <tr>
                    <td>style.colWidth</td>
                    <td>默认列宽</td>
                    <td>默认54</td>
                </tr>
                <tr>
                    <td>style.fontColor</td>
                    <td>字体颜色</td>
                    <td>默认#000000</td>
                </tr>
                <tr>
                    <td>style.fontSize</td>
                    <td>字体大小</td>
                    <td>默认12</td>
                </tr>
                <tr>
                    <td>style.wrapText</td>
                    <td>自动换行</td>
                    <td>默认true</td>
                </tr>
                <tr>
                    <td>style.border</td>
                    <td>object</td>
                    <td>
                        <table>
                            <tr>
                                <td>border.<a href='#borderstyle取值'>style</a></td>
                                <td>边框类型</td>
                                <td>默认无边框</td>
                            </tr>
                            <tr>
                                <td>border.color</td>
                                <td>边框颜色</td>
                                <td>默认#000000</td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>body</td>
        <td>array</td>
        <td>工作表数据列设置</td>
        <td>
            <table>
                <tr>
                    <td>field</td>
                    <td>列字段</td>
                    <td>与data中的字段对应</td>
                <tr>
                    <td>text</td>
                    <td>列标题</td>
                    <td>为空时默认为column[1,2,3,...]</td>
                </tr>
                <tr>
                    <td>width</td>
                    <td>列宽</td>
                    <td>未设置时使用style.colWidth默认列宽</td>
                </tr>
                <tr>
                    <td>formatter</td>
                    <td>格式化函数</td>
                    <td>parameter1:字段值, parameter2:行索引, parameter3:行数据</td>
                </tr>
                <tr>
                    <td>merge</td>
                    <td>合并当前列上相同的数据(相邻行)</td>
                    <td>bool</td>
                </tr>
                <tr>
                    <td>rowspan</td>
                    <td>合并表头(行)</td>
                    <td>number</td>
                </tr>
                <tr>
                    <td>colspan</td>
                    <td>合并表头(列)</td>
                    <td>number</td>
                </tr>
            </table>
        </td>
    </tr>
</table>


#### border.Style取值
- default　　/\*Continuous\*/
- default1
- default2
- default3
- dot
- dash
- dash1
- dashDot
- dashDot1
- dashDots　　/\*dashDotDot\*/
- dashDots1
- slantDashDot
- double

