function checkCapsLock(e)
	{
	var myKeyCode=0;
	var myShiftKey=false;
	var myMsg='Caps Lock is on.\n\nTo prevent entering your password incorrectly,\nyou should press Caps Lock to turn it off.';
	if (document.all)
		{
		myKeyCode=e.keyCode;
		myShiftKey=e.shiftKey;
		}
	else if (document.addEventListener)
		{
		myKeyCode=e.which;
		myShiftKey=e.shiftKey;
		}
	else if (document.layers)
		{
		myKeyCode=e.which;
		myShiftKey=( myKeyCode == 16 ) ? true : false;
		}
	else if (document.getElementById)
		{
		myKeyCode=e.which;
		myShiftKey=( myKeyCode == 16 ) ? true : false;
		}
	if ( ( myKeyCode >= 65 && myKeyCode <= 90 ) && !myShiftKey )
		{
		alert( myMsg );
		}
	else if ( ( myKeyCode >= 97 && myKeyCode <= 122 ) && myShiftKey )
		{
		alert( myMsg );
		}
	}