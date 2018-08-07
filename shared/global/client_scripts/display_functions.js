function MakeUpperCase()
	{
	if(event.keyCode >= 97 && event.keyCode <= 122)
		{
		event.keyCode-=32;
		}
	}

function ShowHideItem(selection, value, id)
	{
	var item = document.all ? document.all(id) :
			   document.getElementById(id);
	if (selection == value)
		{
		item.style.display = '';
		item.style.visibility = 'visible';
		}
	else
		{
		item.style.display = 'none';
		item.style.visibility = 'hidden';
		}
	}

function ChangeDisplay(objItemToChange, objHiddenField)
	{
	if(objItemToChange.style.display=='none')
		{
		objItemToChange.style.display='inline';
		objHiddenField.value='Y';
		}
	else
		{
		objItemToChange.style.display='none';
		objHiddenField.value='N';
		}
	}

function SetDisplay(objItemToSet, objHiddenField)
	{
	if(objHiddenField.value=='Y')
		{
		objItemToSet.style.display='inline';
		}
	else
		{
		objItemToSet.style.display='none';
		}
	}