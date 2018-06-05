export
===
`导出excel表格很简陋 haha`  

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
        <td>每个工作表的名</td>
        <td>如果title.text已设置,则默认title.text,当两者都未设置则为sheet[1,2,...]</td>
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
                    <td>如果sheetName已设置,text未设置,则text默认为sheetName,当两者都未设置则为sheet[1,2,...]</td>
                </tr>
                <tr>
                    <td>title.height</td>
                    <td>标题高度</td>
                    <td>默认50</td>
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
                    <td>style.fontColor</td>
                    <td>字体颜色</td>
                    <td>默认#000000</td>
                </tr>
                <tr>
                    <td>style.border</td>
                    <td>object</td>
                    <td>
                        <table>
                            <tr>
                                <td>border.<a href='#borderStyle'>style</a></td>
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
                    <td></td>
                </tr>
            </table>
        </td>
    </tr>
</table>


# borderStyle
- default   /*Continuous*/
- default1
- default2
- default3
- dot
- dash
- dash1
- dashDot
- dashDot1
- dashDots  /*dashDotDot*/
- dashDots1
- slantDashDot
- double

