var xlsx = require('node-xlsx');  
var fs =require('fs');

var k_v_path = 'src/k-vV2.json';
var k_v_data = fs.readFileSync(k_v_path);
var kv = JSON.parse( k_v_data );

var obj = xlsx.parse('src/杰亚迪 - MTM系统数据表格 2017-8-25中 - 20170912-adj.xlsx'); //当前excel名字  
// console.log(obj[2]);
var clothingTypeDef = {'上衣':'A','马甲':'C','大衣':'E','裤':'B','下装':'B'};
console.log('服装类型----------------',clothingTypeDef);
var factoryClothingModelStyleDefs = [];
// 读取第三到第九个sheet
for (var l = 2;l<10;l++){
		
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
        factoryClothingModelStyleDef.modelVersion = row[6];

        var defaultStyles = [];
        for(var j = 1;j<6;j++){
            if(!row[j])continue;
            var style = {};
            style.name = colNmaes[j];
            if(style.name=='纽扣数') continue;
            if(style.name=='纽扣数N') style.name='纽扣数';
            style.key = kv[style.name];
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
                            if(style.name=='驳宽' && !isNaN(+valueArr[0])) style.value = (+valueArr[0]).toFixed(1);
                            for(var k = 0 ;k<valueArr.length; k++){
                                var option = {};
                                var _value1 = '';
                                if(valueArr[k]){
                                    _value1 = valueArr[k] ? valueArr[k].replace('\n',''):'';
                                }
                                if(style.name=='驳宽'  && !isNaN(+_value1) ) _value1 = (+_value1).toFixed(1);
                                option.shape = _value1;
                                option.size = '';
                                option.imageUrl = '';
                                options.push(option);
                            }
                            
                        }else{
                            var option = {};
                            var _value = valueCell ? valueCell.replace('\n',''):'';
                            if(style.name=='驳宽' && !isNaN(+_value))  _value = (+_value).toFixed(1);
                            option.shape = _value
                            style.value = _value;

                            option.size = '';
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
console.log(factoryClothingModelStyleDefs instanceof Array);
var wPath = 'result/factoryClothingModelStyleDefs'+ (+new Date())+'.json'
fs.writeFile(wPath, JSON.stringify(factoryClothingModelStyleDefs),  function(err) {
   if (err) {
       return console.error(err);
   }
   console.log("数据写入成功！");

});
