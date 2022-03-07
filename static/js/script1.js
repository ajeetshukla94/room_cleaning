var body = $('body');
body.on('click', 'button.add-row', function() 
{
  
  var table = $(this).closest('div.table-content'),
  tbody = table.find('tbody'),
  thead = table.find('thead');

  
  if (tbody.children().length > 0)
  {
	tbody.find('tr:last-child').clone().appendTo(tbody);
  }
  else 
  {
	var trBasic = $('<tr />', 
	{
		'html': '<td><span class="remove remove-row">x</span></td><td><input type="text" class="form-control" /></td>'
	}),
	columns = thead.find('tr:last-child').children().length;
		
	for (var i = 0, stopWhen = columns - trBasic.children.length; i < stopWhen; i++) 
	{
	   $('<td />', {'text': 'static element'}).appendTo(trBasic);
	}
	tbody.append(trBasic);
  }
});

body.on('click', 'span.remove-row', function() 
{
  $(this).closest('tr').remove();
});

body.on('click', 'span.remove-col', function() 
{
  var cell = $(this).closest('th'), 
  index = cell.index() + 1; 
  cell.closest('table').find('th, td').filter(':nth-child(' + index + ')').remove();
});


function submit(){

var data = Array();
code_tbl = document.getElementsByClassName("code_tbl")[0]
var header_length =code_tbl.rows[0].children.length
code_rows = code_tbl.rows
for(var i =0; i<code_rows.length; i++)
{
		
		tds = code_rows[i].children
		data[i] = Array();
		for(var j = 1; j<header_length; j++)
		{
			if (i==0 && j==1)
			{
				data[i][j] = "Equipment"
			}
			else if (i==0 )
			{
				data[i][j] = tds[j].children[2].children[0].firstElementChild.value
			}
			else
			{
				data[i][j]  = tds[j].firstElementChild.value
			}
			
		}
		
}

	$.getJSON('/submit_data', 
	{
		params_data : JSON.stringify(data)
	}, 
	function(result) 
	{
		var link = document.createElement('a')
		link.href =result.file_path;
		link.download = result.file_name;
		link.dispatchEvent(new MouseEvent('click'));
		
	});

	
}