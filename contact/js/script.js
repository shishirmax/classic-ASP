function validate(){
	if(document.form.name.value=="")
	{
		alert("Enter Name")
		document.form.name.focus();
		return false;
	}

	var mailformat = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/; 
	if(document.form.email.value.match(mailformat))
	{
		//continue
	}
	else
	{
		alert("Enter valid Email id like(example@domain.com)");
		document.form.email.focus();
		return false;
	}

	if(document.form.contact.value=="")
	{
		alert("Enter Contact")
		document.form.contact.focus();
		return false;
	}

	if(document.form.address.value=="")
	{
		alert("Enter Address")
		document.form.address.focus();
		return false;
	}

	if(document.form.state.value=="")
	{
		alert("Enter State")
		document.form.state.focus();
		return false;
	}

	if(document.form.city.value=="")
	{
		alert("Enter City")
		document.form.city.focus();
		return false;
	}
	return true;
}