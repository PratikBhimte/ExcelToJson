/*  Controller file  */
/************************************************** */
//Master Controller file 
//Created by : Pratik Bhimte
//Dated : 02-Jan-2014
// Recent Modification : added block to handle special characters
/************************************************** */
/*  Processing Excel  */

var _file_;
var ext;
var JSONBatch=[];
var bn,bt;
var _mainObject_;
var max=5000;
var tolerance = 1000;
var maxRowsToBeDisplayed = 30;
 
var app = angular.module('ExcelExport',['ui.bootstrap']);

var ExcelController = function($scope,$http)
{

  $scope.DispData = {};
  $scope.DataShow = false;
  $scope.loading = false;
  $scope.ExportDone = false;
    
  $scope.loadfileXLSX = function(data)    
  {
	$scope.loading = true;	 
	$('#loader').show();
	
  	var obj = xlsx(data);
  	  	
  	obj = $scope.normalizeJson(obj,'xlsx');
  	  	
    $scope.DispData = obj;
    console.log($scope.DispData);
    _mainObject_ = $scope.DispData;
    $scope.DataShow = true;
    $scope.loading = false;
    $scope.$apply();
    
};


$scope.loadFile = function()
{
 
	JSONBatch=[];
	
	$scope.loading = true;
	
	$('#loader').show();
  var filesSelected = document.getElementById("inputFileToLoad").files;
  if (filesSelected.length > 0) {
  	var fileToLoad = filesSelected[0];

  	var fileReader = new FileReader();

  	fileReader.onload = function (fileLoadedEvent) {
  		if (!$scope.precheck()) {
  			alert('Please select a excel (.xls,.xlsx) file');
  			return;
  		} else {
				
				switch (ext) {
					case "xls":
					_file_ = $scope.makeCompatible(fileLoadedEvent.target.result);
					$scope.loadFileXLS(_file_);
					$scope.ExportShow = true;
					$scope.$apply()
					$scope.loading = false;
					$('#loader').hide();
					break;
					case "xlsx":
					_file_ = $scope.makeCompatible(fileLoadedEvent.target.result);
					$scope.loadfileXLSX(_file_);
					$scope.loading = false;
					$('#loader').hide();
					$scope.ExportShow = true;
					//$scope.pagetotal= Math.ceil($scope.DispData.ExcelData.length/maxRowsToBeDisplayed);
					$scope.$apply()
					break;
				}
			}

		};

		fileReader.readAsDataURL(fileToLoad);
	}

}


$scope.makeCompatible = function(str)
{

	var n = str.indexOf("base64,")
	return str.substr(n + 7, str.length)
}



$scope.precheck = function() 
{
	
	
	var exttemp = document.getElementById('inputFileToLoad').value.split('.');
	ext = exttemp[(exttemp.length-1)]
		
	if (ext.toString().toLowerCase() === 'xlsx' || ext.toString().toLowerCase() === 'xls') {
		return true;
	} else {
		return false;
	}
}

$scope.loadFileXLS = function (data) {
	

	var cfb = XLS.CFB.read(data, {
		type : 'base64'
	});
	var wb = XLS.parse_xlscfb(cfb);
	
	var output = JSON.stringify($scope.to_json(wb), 2, 2)
     
	
	var obj = $.parseJSON(output);

   
   

$scope.DispData = "xls loaded"

$scope.DispData = $scope.normalizeJson(obj,'xls');
_mainObject_ = $scope.DispData;
console.log($scope.DispData);
$scope.DataShow = true;
$scope.loading = false;
     $scope.$apply();
 }


 $scope.to_json = function (workbook) {
 	var result = {};
 	workbook.SheetNames.forEach(function (sheetName) {
 		var roa = XLS.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
 		if (roa.length > 0) {
 			result[sheetName] = roa;
 		}
 	});
 	return result;
 }


 $scope.normalizeJson = function(obj,filetype)
 {

  
  var normalizedJSON = {};
  var columnHeader = [];
  var ExcelData = [];
  var rowData = [];

  if(filetype==='xlsx')
  {
  	var ws = obj.worksheets[0];

       
       var maxRow = ws.maxRow;
       var maxColumn = ws.maxCol;
       var r1;

       var rownum = 0;

       for (var i = 0; i < maxRow; i++) {

       	rowData = [];

       	if (i == 0) {


       		for (var j = 0; j < maxColumn; j++) 
       		{
       			r1 = ws.data[i]

       			if (!r1 || typeof r1 === 'undefined') {
       				break;
       			} else {
       				data = ws.data[i][j];

       				if (!data || typeof data === 'undefined' || typeof data === 'null') {

       					columnHeader.push(null);

       				} else {
       					if (!data.value || typeof data.value == 'undefined' || typeof data.value === 'null') {

       						columnHeader.push("");
       					} else {

       						columnHeader.push(data.value);
       					}
       				}
       			}
       		}

       	} 
       	else {


       		for (var j = 0; j < maxColumn; j++) 
       		{
       			r1 = ws.data[i]

       			if (!r1 || typeof r1 == 'undefined') {
       				break;
       			} else {
       				data = ws.data[i][j];

       				if (!data || typeof data === 'undefined') {

       					rowData.push(null)
       				} else {
       					if (!data.value || typeof data.value === 'undefined' || data.value == '') {
                              
       						rowData.push('');
       					} else {

       						rowData.push(data.value);
       					}
       				}
       			}
       		}
                  
       		if(rowData.length == columnHeader.length)
       		{
       		  ExcelData.push(rowData);	
       		}
       		
       	}

       }
     
       ExcelData = $scope.removeBlankRows(ExcelData,columnHeader.length);
       normalizedJSON.columnHeader = columnHeader;
       normalizedJSON.ExcelData = ExcelData;
        return normalizedJSON;

   }


   if(filetype ==='xls')
   {

     

var newobj = $scope.getSamtolJSON(obj)
var datalenght = newobj.length-1;
  normalizedJSON.columnHeader = newobj.splice(0,1)[0];
  normalizedJSON.ExcelData = newobj.splice(0,datalenght)
  
  
return normalizedJSON;
    }



}


$scope.removeBlankRows = function(rowData,columnCount)
{
 var rdata;
 var cnt;
 var trimmedRowData =[]
for(var i=0;i<rowData.length;i++)
{
 rdata=[];
 rdata = rowData[i];
 
 if(rdata.length != columnCount)
	 {
	 
	 console.log('length mismatch for index =' + i)
	 }
cnt = 0;
  for(var j=0;j<rdata.length;j++)
	  {
	     if(rdata[j]==''|| !rdata[j] || rdata[j] === null || rdata[j] === undefined || rdata[j].length == 0 )
	    	{
	        cnt=cnt+1;	 
	    	}
	   }
 if(cnt >= columnCount)
	 {
	   
	 }
 else
	 {
	 trimmedRowData.push(rdata);
	 cnt = 0;
	 }
 
}
   	
 return trimmedRowData;
	
}
 
 
 
$scope.getSamtolJSON = function(obj)
{

	var maxIndex,maxnum;
	num1= -1;
	num2 = -1;

	var lengthArr = [];
	var columnHeader = [];

	var ReformatedData = [];
	var rowData = [];
	var rowCol = [];

	for(var i =0;i<obj.Sheet0.length;i++)
	{
		lengthArr.push(Object.keys(obj.Sheet0[i]).length);
	}

maxnum = Math.max.apply(Math,lengthArr);
maxIndex = lengthArr.indexOf(Math.max.apply(Math,lengthArr));

$.each(obj.Sheet0[maxIndex],function(key,val){

	columnHeader.push(key);

})

ReformatedData.push(columnHeader);

var flag;
var datavalue;
for(var i=0;i<obj.Sheet0.length;i++)
{
	rowData = [];
	flag = false;
	$.each(columnHeader,function(colkey,colvalue)
	{
        
        flag=false;
         datavalue="";
        
		$.each(obj.Sheet0[i],function(key,value)
		{
                    
			if(colvalue == key)
			{
				flag=true;
				datavalue= value;
				
				return;
			}
         
		});

		
   if(flag===true)
		{
			
			rowData.push(datavalue)
		}
		else
		{
			rowData.push('')
		}


	});

	ReformatedData.push(rowData);


}

return ReformatedData;

}


$scope.makePostReady = function(obj)
{
var ReadyToBeSent = {};
var modifiedColumnHeader=[];
var modifiedExcelData = [];
var modifiedRowData = [];
var modifiedHolder = []
for(var i=0;i<obj.columnHeader.length;i++)
{
  modifiedColumnHeader.push($scope.EscapeUmlaute(obj.columnHeader[i])) 	
}

for(var i=0;i<obj.ExcelData.length;i++)
{
  modifiedRowData = [];
  modifiedRowData = obj.ExcelData[i];
  modifiedHolder = [];
  for(var j=0;j<modifiedRowData.length;j++)
	  {
	  
	    modifiedHolder.push($scope.EscapeUmlaute(modifiedRowData[j]))  
	  }

  modifiedExcelData.push(modifiedHolder);
}

ReadyToBeSent.columnHeader = modifiedColumnHeader;
ReadyToBeSent.ExcelData = modifiedExcelData;

return ReadyToBeSent;

}

$scope.EscapeUmlaute = function(instring)
{   
	var Ismatch = null;
	//console.log(instring)
	
	if(instring === null)
		{
			return ""
		}
	
	if(!instring || typeof instring!='undefined' || instring != null)
		{
		Ismatch = instring.toString().match(/[ÿþýüûúùø÷öõôóòñðïîíìëêéèçæåäãâáàßÞÝÜÛÚÙØ×ÖÕÔÓÒÑÐÏÎÍÌËÊÉÈÇÆÅÄÃÂÁÀ¿¾½¼»º¹¸·¶µ´³²±°¯®¬«ª©¨§¦¥¤¡¢£]/gi);	
		
		}
   	
	
	if(!Ismatch || Ismatch === null || Ismatch.length==0)
		{
		  return instring
		}
	else
		{
		   
		return escape(instring);		
		}
	
}



$scope.CreateBatch = function(RawExcelData,columnCount)
{
	   var k = "20";
	      $scope.progress= k.toString();
		    $scope.progressText="Creating batch ..."
			//$scope.$apply();
	
	    
	var N = ((RawExcelData.length)*(columnCount));
	var cnt,len,diff;
		
	if(!RawExcelData) {return}
	
	if(!JSONBatch || JSONBatch.length == 0)
	{
	  	cnt=0;
	  	len=0;
	  	diff=0;
	}
	else
	{
		cnt = JSONBatch.length;
		len=0;
		diff=0;
	}
	
	if(N > (max+tolerance))
	{
	  
		len = Math.round(max/columnCount); 
	   JSONBatch.push(RawExcelData.splice(0,len));
	  diff = ((N)-(max));
	  if(diff > (max+tolerance))
		{
		  
		  $scope.CreateBatch(RawExcelData,columnCount);
		  return;
		}
	  else
		  {
		  JSONBatch.push(RawExcelData.splice(0,RawExcelData.length));
		  //$scope.CreateBatch(RawExcelData,columnCount);
		  return;
		  
		  }
	}
	else
	{
		JSONBatch.push(RawExcelData);
		return true;
				
	}
	
}

$scope.makeJSONoutOfBatch = function(BatchData,columnData)
{
	var bhejDo={}
	bhejDo.columnHeader = columnData;
	bhejDo.ExcelData = BatchData;
	
	return bhejDo;
	
}




}



