export
===`导出excel表格很简陋 haha`  

Docs
---
**用法**
```javascript
xyz_export.sheet({
    fileName:"计划导出",
    data:[[a:1,b:2,c:3],...],
    sheets:[
        {
            sheetName:"sheet1",
            title: {
                text: "xx月xx日计划",
                height: 50
            },
            style:{
                headHeight:30,
                rowHeight:15,
                fontColor:"",
                border:{
                    style:"default",
                    color:"#000000"
                }
            },
            body: [
                { field: "a", text: "字段1", width: 50 },
                { field: "b", text: "字段2", width: 50 },
                { field: "c", text: "字段3", width: 50 }
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

##### sheet
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
        <td>每个工作表的名</td>
        <td>如果title.text已设置,则默认title.text,当两者都未设置则为sheet[1,2,...]</td>
    </tr>
    <tr>
        <td>title</td>
        <td>object</td>
        <td>工作表标题设置</td>
        <td>不设置title时,默认不显示sheet标题
```
text:标题名字       如果sheetName已设置,text未设置,则text默认为sheetName,当两者都未设置则为sheet[1,2,...]
height:标题高度     默认50
```
        </td>
    </tr>
    <tr>
        <td>style</td>
        <td>object</td>
        <td>工作表样式设置</td>
        <td>
```
headHeight:表头高度     默认15
rowHeight:行高        默认13.5
fontColor:字体颜色      默认#000000
border:{
            style:边框类型      默认无边框
            color:边框颜色      默认#000000
            }
```
        </td>
    </tr>
    <tr>
        <td>body</td>
        <td>array</td>
        <td>工作表数据列设置</td>
        <td>
```
field:列字段,和data中的字段对应
text:列标题    为空时默认为column[1,2,3,...]
width:列宽
```
        </td>
    </tr>
</table>




