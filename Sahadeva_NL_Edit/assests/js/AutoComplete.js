// JScript File

    function ListSelected(source, eventArgs) 
    {
        var editableElement = source.get_element();
        
        var x = editableElement.id;
        if(document.getElementById(x.substring(0,x.lastIndexOf("_")) + "_hidClientItemSelectedFunction").value != "")
        {
            if(!eval(document.getElementById(x.substring(0,x.lastIndexOf("_")) + "_hidClientItemSelectedFunction").value))
            {
                eventArgs.IsValid = false;
                return false;
            }
        }
      
        document.getElementById(x.substring(0,x.lastIndexOf("_")) + "_hidSearchValue").value = eventArgs.get_value();
        document.getElementById(x.substring(0,x.lastIndexOf("_")) + "_hidOrgVal").value = eventArgs.get_value();
        document.getElementById(x.substring(0,x.lastIndexOf("_")) + "_hidOrgText").value = eventArgs.get_text();
        RaiseChangeEvent(x.substring(0,x.lastIndexOf("_")));
    }
    function TextChanged(hidSearchValue, txtSearch, hidOrgVal, hidOrgText)
    {
        document.getElementById(hidSearchValue).value = "";
        if(document.getElementById(txtSearch).value == "")
        {
            document.getElementById(hidOrgVal).value = "";
            document.getElementById(hidOrgText).value = "";
        }
    }
    function TextBlur(hidSearchValue, txtSearch, hidOrgVal, hidOrgText)
    {
        if(document.getElementById(hidSearchValue).value == "")
        {
            if(document.getElementById(txtSearch).value != "")
            {
                document.getElementById(hidSearchValue).value = document.getElementById(hidOrgVal).value;
                document.getElementById(txtSearch).value = document.getElementById(hidOrgText).value;
            }
            else
                document.getElementById(txtSearch).value = "";
        }
    }
    function TextFocus(hidSearchValue, txtSearch, hidOrgVal, hidOrgText)
    {
        document.getElementById(hidOrgVal).value = document.getElementById(hidSearchValue).value;
        document.getElementById(hidOrgText).value = document.getElementById(txtSearch).value;
    }
    
    function RaiseChangeEvent(ctrlID)
    {
        if(document.getElementById(ctrlID + "_hidRaiseEvent").value == "1")
        {   
            __doPostBack(ctrlID, '');
        }
    }
