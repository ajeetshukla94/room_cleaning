class MyHeader extends HTMLElement{
	
	connectedCallback()	{
		this.innerHTML ='<div id="header"> \
		<h2 id="header-text">Room Cleaning Validation  Solution</h2> \
		</div> \
		<div class="navbar"> \
		<a href="/cleaning_room">Cleaning Room</a>  \
		<a href="/UpdateProductList">Update Product List</a>  \
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