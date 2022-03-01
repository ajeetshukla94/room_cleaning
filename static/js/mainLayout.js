class MyHeader extends HTMLElement{
	
	connectedCallback()	{
		this.innerHTML ='<div id="header"> \
		<img src="static/images/ppe.png" id="header-img"> \
		<h2 id="header-text">HVAC SOLUTION</h2> \
		</div> \
		<div class="navbar"> \
		<a href="/logout">LOGOUT</a>  \
		</div>\
		'
	}
}
customElements.define('my-header',MyHeader)


class MyFooter extends HTMLElement{
	
	connectedCallback()	{
		this.innerHTML ='<div id="footer">\
			<h6 id="footer-text">Copyright &#169; Pin Point Engineers</h6>\
		</div>\
		'
	}
}
customElements.define('my-footer',MyFooter)