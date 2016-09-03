function AppendNote()
{
    var dt = new Date();
    Xrm.Page.getAttribute("description")
            .setValue("Update in" + dt.toString());
}