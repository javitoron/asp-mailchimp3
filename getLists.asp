<!--#include file="lib/BF_MC_list.asp"-->
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <body>
        <% call mc.getRequest("lists", null, null) %>
        <h1>Listas</h1>
        <%
            dim myJSON
            set myJSON = new aspJSON
            myJSON.loadJSON_from_string(mc.getLastResponseBody())

            dim list, this_list
            for each list in myJSON.data("lists")
                Set this_list = myJSON.data("lists").item(list)
                Response.Write this_list.item("id") & ": " & _
                    URLdecode( this_list.item("name") ) & "<br>"

            next
        %>
        <h1>Response</h1>
        <%
            Response.Write myJSON.JSONoutput()
        %>
    </body>
</html>