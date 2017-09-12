var xlsx = require('node-xlsx');  
var fs =require('fs');

var k_v_path = 'src/k-v.json';
var k_v_data = fs.readFileSync(k_v_path);
var kv = JSON.parse( k_v_data );

var obj = xlsx.parse('src/MTM系统数据表格2017-8-25中.xlsx'); //当前excel名字  
// console.log(obj[2]);
var clothingTypeDef = {'上衣':'A','马甲':'C','大衣':'E','裤':'B','下装':'B'};
console.log('服装类型----------------',clothingTypeDef);
var factoryClothingModelStyleDefs = [];
for (var l = 2;l<obj.length;l++){
		
	var sheet = obj[l];
    var name = sheet.name;//sheet name
    var genderName = name.substring(0, 1);
    var gender = genderName==='男' ? 1:0;
    var clothingTypeName = name.substring(1);
    var clothingType = clothingTypeDef[clothingTypeName];
    var factoryNumber = 'ABBO';
    console.log(name);

    var colNmaes = sheet.data[0];//表头
	for (var i = 1 ; i < sheet.data.length-1; i++) {
		var factoryClothingModelStyleDef = {};
		var row = sheet.data[i];
        if(!row[0]) continue;
        
        // console.log('colNmaes--------------',colNmaes);
		// console.log(row);
        factoryClothingModelStyleDef._class='cn.com.icaifeng.model.style.FactoryClothingModelStyleDef'
		factoryClothingModelStyleDef.factoryNumber = factoryNumber;
        factoryClothingModelStyleDef.gender = gender;
		factoryClothingModelStyleDef.specId = '';

		factoryClothingModelStyleDef.clothingType = clothingType;
        var modelImagesUrl = 'modelImages/'+factoryNumber+'/'+clothingType+'/'+gender+'/'+row[0]+'.jpg';
        var modelImages = [];
        modelImages.push(modelImagesUrl);
		factoryClothingModelStyleDef.modelImages = modelImages;
        factoryClothingModelStyleDef.modelVersion = row[5];

        var defaultStyles = [];
        for(var j = 1;j<5;j++){
            if(!row[j])continue;
            var style = {};
            style.name = colNmaes[j];
            style.key = kv[colNmaes[j]];
            style.required = '0'
            style.order = j - 1;
            style.variable = false;
            var options = [];
            var valueCell = row[j]+'';
            var valueArr = [];
            if(valueCell){
                if(typeof valueCell ==='string'){
                    valueArr = valueCell.split('/');
                    if(valueArr){
                        if(valueArr.length>1){
                            style.variable = true;
                            style.value = valueArr[0];
                            for(var k = 0 ;k<valueArr.length; k++){
                                var option = {};
                                var _value = '';
                                if(valueArr[k]){
                                    _value = valueArr[k] ? valueArr[k].replace('\n',''):'';
                                }
                                option.shape = '';
                                option.size = '0cm';
                                option.imageUrl = '';
                                options.push(option);
                            }
                            
                        }else{
                            var option = {};
                            var _value = valueCell ? valueCell.replace('\n',''):'';
                            option.shape = _value
                            style.value = _value;
                            option.size = '0cm';
                            option.imageUrl = '';
                            options.push(option);
                            if(option.shape=='undefined')
                            console.log(option.shape,j,row[j]);
                        }
                    }
                }

                
            }

            style.options = options;

            defaultStyles.push(style);
        }
        

		factoryClothingModelStyleDef.defaultStyles = defaultStyles;
		// console.log(factoryClothingModelStyleDef);
		factoryClothingModelStyleDefs.push(factoryClothingModelStyleDef);
	}
	
}
// console.log(factoryClothingModelStyleDefs);
var wPath = 'result/factoryClothingModelStyleDefs.json'
fs.writeFile(wPath, JSON.stringify(factoryClothingModelStyleDefs),  function(err) {
   if (err) {
       return console.error(err);
   }
   console.log("数据写入成功！");

});

// console.log(demos);
//写入第二个sheet
// const data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
// var buffer = xlsx.build([{name: "Sheet4", data: data}]);
// console.log(buffer.toString());
// fs.writeFileSync('result/savepath.xlsx',buffer,{'flag':'w'});
// for(var i=0;i<obj[0].data.length;i++)  
//     for(var j=0;j<obj[0].data[0].length;j++)  
// 		console.log(obj[0].data[i][j]); 