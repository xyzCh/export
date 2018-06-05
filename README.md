export
===
`导出excel表格很简陋 haha`  

Docs
---
***
**用法**
```
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

