function close_info(){
	document.getElementById('alert_info').style.visibility = "hidden"
}

function close_err(){
	document.getElementById('alert_error').style.visibility = "hidden"
}

function on() {
  document.getElementById("overlay").style.display = "block";
  
}

function off() {	
  document.getElementById("overlay").style.display = "none";
}