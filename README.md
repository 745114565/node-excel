# node-excel
read and write data in excel use node 
# dev step
* install nodeJs
* install node-xlsx

``` bash
  npm install node-xlsx 
```
或者

``` Bash
  npm install -g node-xlsx 
```
或者

``` Bash
  npm install node-xlsx --save-dev
```

* 加载excel模块

``` javascript
  var xlsx = require("node-xlsx");
```

* 读取excel

``` javascript
var list = xlsx.parse("./excel/" + excelName);
```

* 读出后是数组，包含每个sheet

``` txt
  [

      { name: 'sheet1',data: [ [Object], [Object], [Object], [Object], [Object] ] },
      { name: 'sheet2', data: [ [Object] ] }

  ]
```
  name=sheet名称

  data=每个sheet的数据，

  剩下的就灵活操作咯......
