var totalCellDataToSave = {};
var totalCellDataGot = {};

/**
 * This function is called whenever add row button is clicked
**/
function addRow()
{
	var table = document.getElementById("excel-table");
	
	var lastRow = document.getElementById("excel-table").rows[document.getElementById("excel-table").rows.length-1].id;
	
	var columnCount = document.getElementById("excel-table").rows[document.getElementById("excel-table").rows.length-1].children.length;

	var newRowNumber = parseInt(lastRow.replace("row",""))+1;
	
	var newRow = table.insertRow();

	newRow.className = "excel-row";
	
	newRow.id = "row"+newRowNumber;
	for(var i=0;i<columnCount;i++){
		var newCol = document.getElementById("row"+newRowNumber).insertCell();
		newCol.id  = "data"+String(newRowNumber)+String(i);
		newCol.setAttribute("contentEditable",true);
 	}

	makeCellsEditable();

	attachEventListenersToCells()
}

/**
 * This function is called whenever delete row button is clicked
**/
function deleteRow()
{
	var table = document.getElementById("excel-table");

	var lastRowNumber = document.getElementById("excel-table").rows.length-1;

	table.deleteRow(lastRowNumber);
}

/**
 * This function is called whenever add column button is clicked
**/
function addColumn()
{
	var rowsCount = document.getElementById("excel-table").rows.length;
	var lastHeading = document.getElementById("row1").lastElementChild.innerHTML;
	var columnCount = document.getElementById("excel-table").rows[document.getElementById("excel-table").rows.length-1].children.length;
	var nextHeading = nextChar(lastHeading);
	if(typeof nextHeading != "undefined"){
		document.getElementById("row1").insertCell().outerHTML = "<th>"+nextHeading+"</th>";
		for (var i=1;i<rowsCount;i++)
		{
			var newCell = document.getElementById("row"+(i+1)).insertCell();
			newCell.id  = "data"+String(i)+String(columnCount);
			newCell.setAttribute("contentEditable",true);	
		}
	}

	makeCellsEditable();

	attachEventListenersToCells()

}

/**
 * This function is called whenever delete column button is clicked
**/
function deleteColumn()
{
	var rowsCount = document.getElementById("excel-table").rows.length;
	
	for(var i=0;i<rowsCount;i++)
	{
		document.getElementById("row"+(i+1)).deleteCell(-1);
	}

}


/**
 * This function gives the next character in the alphabetic sequence till Z
 * i.e if we give 'A' as input it returns 'B' as output
**/
function nextChar(c) {
	if(c.length == 1 && c!='Z'){
    	return String.fromCharCode(c.charCodeAt(0) + 1);
	}else{
		
	}
}


/**
 * This function makes all the data cells which are available "editable"
**/
function makeCellsEditable()
{
	for(var i=0;i<document.getElementsByTagName('td').length;i++){
		document.getElementsByTagName('td')[i].setAttribute("contentEditable",true);		
	}
}

/**
 * Attach event listeners to all the cells. They are useful to save the data entered by the user
**/

function attachEventListenersToCells()
{
		for(var i=0;i<document.getElementsByTagName('td').length;i++){
			document.getElementsByTagName('td')[i].addEventListener('click',function(){
			});	
			document.getElementsByTagName('td')[i].addEventListener('focus',function(){
			});	
			document.getElementsByTagName('td')[i].addEventListener('focusout',function(){
				setDataInLocalStorage();
				getDataFromLocalStorage();
			});	
		}	
}

/**
 * Attach event listeners to all the cells. They are useful to save the data entered by the user
**/
function setDataInLocalStorage()
{
	var totalCellDataGot = JSON.parse(localStorage.ExcelSheetData);
	for(var i=0;i<document.getElementsByTagName('td').length;i++)
	{
		if(document.getElementsByTagName('td')[i].innerHTML!=""){
			totalCellDataGot[document.getElementsByTagName('td')[i].id] = document.getElementsByTagName('td')[i].innerHTML;
		}
	}
	localStorage.setItem("ExcelSheetData", JSON.stringify(totalCellDataGot));

}

/**
 * Attach event listeners to all the cells. They are useful to save the data entered by the user
**/
function getDataFromLocalStorage()
{
	totalCellDataGot = 	JSON.parse(localStorage.ExcelSheetData);
	for (var prop in totalCellDataGot) {
		document.getElementById(prop).innerHTML = totalCellDataGot[prop]; 
	}
	console.log(totalCellDataGot);
}


document.addEventListener("DOMContentLoaded", function(){

	//whenever the page is loaded calling this function makes all the available data cells editable
	makeCellsEditable();

	attachEventListenersToCells();

	//call this function to set data in local storage
	setDataInLocalStorage();

	//call this function to get data from local storage and populate in the cells
	getDataFromLocalStorage();

	$(document).keypress("b",function(e) {
  		if(e.keyCode==2 && e.ctrlKey){
  			console.log(this);
  			for(i=0;i<document.getElementsByTagName('td').length;i++){
  				if($($('td')[i]).hasClass('bold-text')){
  					$($('td')[i]).removeClass('bold-text');
  				}else{
  					  document.getElementsByTagName('td')[i].className += ' bold-text';
  				}
  			}
  		}
	});

	$(document).keypress("i",function(e) {
  		if(e.keyCode==9 && e.ctrlKey){
  			console.log(this);
  			for(i=0;i<document.getElementsByTagName('td').length;i++){
  				if($($('td')[i]).hasClass('italic-text')){
  					$($('td')[i]).removeClass('italic-text');
  				}else{
  					  document.getElementsByTagName('td')[i].className += ' italic-text';
  				}
  			}
  		}
	});


	$(document).keypress("u",function(e) {
  		if(e.keyCode==21 && e.ctrlKey){
  			for(i=0;i<document.getElementsByTagName('td').length;i++){
  				if($($('td')[i]).hasClass('underlined-text')){
  					$($('td')[i]).removeClass('underlined-text')
  				}else{
  					  document.getElementsByTagName('td')[i].className += ' underlined-text';
  				}
  			}  			 
  		}
	});
		
});











