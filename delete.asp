<!--#include file="lib/BF_MC_list.asp"-->
<%
    dim list_id, apikey
    list_id = "0b43e32f1a"
    dim subscriber_hash    
    subscriber_hash = mc.subscriberHash(Request.Form("email"))

    'response.write subscriber_hash

    call mc.delete("lists/" & list_id & "/members/" & subscriber_hash, null, null)
    
    if mc.success() then
        response.write "<p>Suscriptor Mailchimp eliminado</p>"
    else
        response.write "<p>Error al eliminar el suscriptor en Mailchimp</p>"
     '   echo "<p>" . $MailChimp->getLastError() . "</p>";
    end if
%>