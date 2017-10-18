function ExportsTest(){
	this.aa = "b";
}

var exportsTest = new ExportsTest();

console.log(exportsTest.aa);