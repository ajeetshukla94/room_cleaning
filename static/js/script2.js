var frame_count = 1;
function add_row(elem){
	
	var row_index = elem.parentElement.parentElement.children[0].rows.length - 1;
	var row = elem.parentElement.parentElement.children[0].rows[row_index];
    var table = elem.parentElement.parentElement.children[0];
    var clone = row.cloneNode(true);
	var cells_length = clone.cells.length - 1;
	for(var i = 1; i<cells_length; i++){
		clone.cells[i].children[0].value=""
	}
    table.appendChild(clone);
}

function delete_row(elem){
	if(elem.parentElement.parentElement.parentElement.rows.length > 2){
		elem.parentElement.parentElement.remove();
	}else{
		alert("First row can not be deleted!")
	}
}


function submit()
{
	code_tbl = document.getElementsByClassName("code_tbl")[0]
	code_rows = code_tbl.rows
	var final_table_data = {};
    var full_data = {};	
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children	
		var table_data = {};
		table_data['Product_Name'] =tds[0].firstElementChild.value 	
		table_data['Generic_Name'] =tds[1].firstElementChild.value 	
		table_data['Form'] =tds[2].firstElementChild.value 	
		table_data['API_with_strength'] =tds[3].firstElementChild.value 	
		table_data['Minimum_Batch_size'] =tds[4].firstElementChild.value 	
		table_data['MRDD'] =tds[5].firstElementChild.value 	
		table_data['LD50'] =tds[6].firstElementChild.value
		table_data['NOEL'] =tds[7].firstElementChild.value		
		
		final_table_data[j] = table_data
	}
	
	full_data['observation']=final_table_data
	
	
	$.getJSON('/submit_UpdateProductList', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		alert("done");		
	});
	
}

