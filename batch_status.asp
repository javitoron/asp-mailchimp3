<!--#include file="lib/setup_MC.asp"-->
<%
    dim bat
    set bat = new MCBatch

    call bat.init(mc, null)


    dim status
    status = bat.check_status("f1183dc938")
    response.write "<h1>Batch status</h1>"
    response.write "<h2>JSON Response</h2>"
    response.write status

    dim myJSON
    set myJSON = new aspJSON
    myJSON.loadJSON_from_string(status)

    response.write "<h2>Summary</h2>"

    Response.Write "finished_operations: " & _
        myJSON.data("finished_operations") & "<br>" & _
        "errored_operations: " & _
        myJSON.data("errored_operations") & "<br><br>"

    response.write "<h2>URLs</h2>"
    dim link, this_link
    for each link in myJSON.data("_links")
        Set this_link = myJSON.data("_links").item(link)
        Response.Write this_link.item("method") & ": " & _
            this_link.item("href") & "<br>"
    next
    
%>