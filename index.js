var xlsx = require('node-xlsx');  
var fs =require('fs');


var obj = xlsx.parse('src/demo.xlsx'); //当前excel名字  
// console.log(obj);

var demos = [];
for (var j = 0;j<obj.length;j++){
		
	var sheet = obj[j];
	for (var i = 0 ; i < sheet.data.length; i++) {
		var demo = {};
		var row = sheet.data[i];
		// console.log(row);
		// console.log(row[0],row[1],row[2],row[3],row[4]);
		demo.hello = row[0];
		demo.world = row[1];
		demo.age = row[2];
		demo.name = row[3];
		demo.count = row[4];
		console.log(demo);
		demos.push(demo);
	}
	
}
console.log(demos);
//写入第二个sheet
const data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
var buffer = xlsx.build([{name: "Sheet4", data: data}]);
// console.log(buffer.toString());
fs.writeFileSync('result/savepath.xlsx',buffer,{'flag':'w'});
// for(var i=0;i<obj[0].data.length;i++)  
//     for(var j=0;j<obj[0].data[0].length;j++)  
// 		console.log(obj[0].data[i][j]); 