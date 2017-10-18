var NodeExcel = require ("./index.js");

var excel = new NodeExcel([[1,2,3,],[4,5,6],[11,12,13]]);
var newFile = excel.transposeData(0,0,3,3);
excel.createNewFile("new", "textNew", "./", newFile);