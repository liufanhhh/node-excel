'user strict';

var xlsx = require('node-xlsx');
var fs = require('fs');

function NodeExcel(file, sheetName = null){
	if (typeof(file) == "string") {
		var path = file;
		this.oldSheets = xlsx.parse(path);
		this.sheet = this._init(this.oldSheets, sheetName);	
	} else if (file && file[0] && file[0][0]){
		this.sheet = file;
	}	
}

NodeExcel.prototype._init = function(oldSheets, sheetName) {
	if (sheetName) {
		for (let i = oldSheets.length - 1; i >= 0; i--) {
			if (oldSheets[i].name&&oldSheets[i].name == sheetName) {
				return oldSheets[i].data;
			}
		}
	} else{
		return oldSheets[0];
	};
};

NodeExcel.prototype._transferColumn = function(columnNumber) {
	var transferedColumn;
	if (columnNumber === null) {
		return false;
	}

	if ( (typeof(columnNumber) !== "number"&&columnNumber.length>3) || (typeof(columnNumber) === "number" && columnNumber>16384) ){
		console.log(" Error! Larger than maximum column number (XFD 16384)");
		return false;
	}

	if(typeof(columnNumber) !== "number"){
		columnNumber = columnNumber.toUpperCase();	
		if (columnNumber[2]) {
			transferedColumn = (parseInt(columnNumber.charCodeAt(0))-64)*26*26+(parseInt(columnNumber.charCodeAt(1))-64)*26+parseInt(columnNumber.charCodeAt(2))-64;
		} else if (columnNumber[1]){
			transferedColumn = (parseInt(columnNumber.charCodeAt(0))-64)*26+parseInt(columnNumber.charCodeAt(1))-64;
		} else if (columnNumber[0]){
			transferedColumn = parseInt(columnNumber.charCodeAt(1))-64;
		}
		return transferedColumn;
	}
	return columnNumber;
}

NodeExcel.prototype.getData = function(columnFrom, rawFrom, columnTo = null , rawTo = null) {
	var columnFrom = this._transferColumn(columnFrom);
	var columnTo = this._transferColumn(columnTo);
	var newTable = [];

	if (columnTo == null) {
		return this.sheet[rawFrom][columnFrom];
	}

	if (columnFrom&&columnTo){
		for (let i = 0; i < rawTo - rawFrom; i++) {
			newTable[i] = [];
			for (let j = 0; j < columnTo - columnFrom; j++) {
				newTable[i][j] = this.sheet[rawFrom+i][columnFrom+j];
			}
		}
	return newTable;
	}

}

NodeExcel.prototype.replaceData = function(data, columnFrom, rawFrom, columnTo = null , rawTo = null) {
	var columnFrom = this._transferColumn(columnFrom);
	var columnTo = this._transferColumn(columnTo);

	if (columnTo == null) {
		return this.sheet[raw].splice(columnTo, data.length, ...data);
	}

	if (columnFrom && columnTo){
		for (let i = 0; i < rawTo - rawFrom; i++) {
			this.sheet[raw].splice(columnTo, data.length, ...data[i]);
		}
		return this.sheet;
	}

}

NodeExcel.prototype.transposeData = function(columnFrom, rawFrom, columnTo = null , rawTo = null ) {
	var columnFrom = this._transferColumn(columnFrom);
	var columnTo = this._transferColumn(columnTo);
	var newTable = [];
	if (columnFrom!==false && columnTo!==false){
		for (let i = 0; i < columnTo - columnFrom; i++) {
			newTable[i] = [];
			for (let j = 0; j < rawTo - rawFrom; j++) {
				newTable[i][j] = this.sheet[rawFrom+j][columnFrom+i];
			}
		}		
	}
	return newTable;
}

NodeExcel.prototype.createNewFile = function(fileName, sheetName, path, data) {
	var buffer = xlsx.build([{name: sheetName, data: data}]);

	fs.writeFile(path+fileName, buffer, function(error){
		if (error) {
			console.log(error);
		} else {
			console.log("create file successful");
		}
	})
}

NodeExcel.prototype.getSingleSheet = function(sheetName) {
	return this._init(this.oldSheets, sheetName);	
}

NodeExcel.prototype.getAllSheets = function() {
	return this.oldSheets;
}

module.exports = NodeExcel;

