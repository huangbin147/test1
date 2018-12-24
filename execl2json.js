var xlsx = require("node-xlsx");
var fs = require('fs');

var list = xlsx.parse('E:/myproject/e2j/routes/data/practice_point_player.xlsx');

praseExcel(list);

function praseExcel(list) {
    console.log(list[0].data);

        var excleData = list[0].data;
        var sheetArray = [];
        var typeArray = excleData[1];
        var keyArray = excleData[2];

        for (var j = 4; j < excleData.length; j++) {
            var curData = excleData[j];
            if (curData.length == 0) continue;
            var item = changeObj(curData, typeArray, keyArray);

            sheetArray.push(item);
        }
        if (sheetArray.length > 0) writeFile("practice_point_player" + ".json", JSON.stringify(sheetArray));

}

//转换数据类型
function changeObj(curData, typeArray, keyArray) {
    var obj = {};

    for (var i = 0; i < keyArray.length; i++) {
        //字母
        obj[keyArray[i]] = changeValue(curData[i], typeArray[i]);
    }
    return obj;
}

function changeValue(value, type){
    if (value == undefined &&type=="string"){
            return "";
    }
    if(value == undefined && type=="int"||value == undefined && type=="number"){
           return 0.00;
    }
    //if(value==undefined) return "";
    if(type=="number"||type=="int")return value;
    if (type == "String"||type=="string") return value;
}

//写文件


function writeFile(fileName, data) {
    fs.writeFile(fileName, data, 'utf-8', complete);

    function complete(err) {
        if (!err) {
            console.log("文件生成成功");
        }
    }
}
