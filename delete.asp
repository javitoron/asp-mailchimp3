<!--#include file="lib/setup_MC.asp"-->
<%
    dim list_id
    list_id = "0b43e32f1a"
    dim subscriber_hash    
    subscriber_hash = mc.subscriberHash(Request.Form("email"))

    call mc.delete("lists/" & list_id & "/members/" & subscriber_hash, null, null)
    
    if mc.success() then
        response.write "<p>Deleted Mailchimp subscriber</p>"
    else
        response.write "<p>Error deleting Mailchimp subscriber</p>"
        
        dim myJSON
        set myJSON = new aspJSON
        myJSON.loadJSON_from_string(mc.getLastResponseBody())

        Response.Write "<p>" & _
            "<strong>" & myJSON.data.item("title") & "</strong>: " & _
            myJSON.data.item("detail") & "</p>"
    end if
%>
