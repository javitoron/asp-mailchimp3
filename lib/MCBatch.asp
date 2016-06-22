<%
 ' A MailChimp Batch operation.
 ' http://developer.mailchimp.com/documentation/mailchimp/reference/batches/
 ' Based on PHP version: https://github.com/drewm/mailchimp-api
 '
 ' @author Javier Torón <javitoron@gmail.com>
 ' @version 0.1

class MCBatch

    private mc

    private operations, ops
    private batch_id


    Private Sub Class_Initialize()
        set operations = New aspJSON
        operations.data.add "operations", operations.Collection()
        ops = 0
    End Sub


    Private Sub Class_Terminate()
    End Sub


    public sub init(p_mc, p_batch_id)
        set mc = p_mc
        batch_id = p_batch_id
    end sub


    public function getOperationsString()
        getOperationsString = operations.JSONoutput()
    end function

    
     'Add an HTTP DELETE request operation to the batch - for deleting data
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
    public sub delete(id, method)
        call queueOperation("DELETE", id, method)
    end sub

    
     'Add an HTTP GET request operation to the batch - for retrieving data
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
     '@param   JSON arguments (usually your data)
    public sub getRequest(id, method, args)
        call queueOperation("GET", id, method, args)
    end sub

    
     'Add an HTTP PATCH request operation to the batch - for performing partial updates
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
     '@param   JSON arguments (usually your data)
    public sub patch(id, method, args)
        call queueOperation("PATCH", id, method, args)
    end sub

    
     'Add an HTTP POST request operation to the batch - for creating and updating items
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
     '@param   JSON arguments (usually your data)
    public sub post(id, method, args)
        call queueOperation("POST", id, method, args)
    end sub

    
     'Add an HTTP PUT request operation to the batch - for creating new items
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
     '@param   JSON arguments (usually your data)
     public sub putRequest(id, method, args)
        call queueOperation("PUT", id, method, args)
    end sub

    
     'Execute the batch request
     '@param int timeout Request timeout in seconds (optional)
     '@return  API response string
     public function execute(p_timeout)
        dim s_response
        if IsNull(p_timeout) then
            p_timeout = 10
        end if
        
        dim req
        set req = operations

        call mc.post("batches", req, p_timeout)

        s_response = mc.getLastResponseBody()
        if mc.success then
            dim myJSON
            set myJSON = new aspJSON
            myJSON.loadJSON_from_string( s_response )

            batch_id = myJSON.data("id")
        end if

        execute = s_response
    end function

    
     'Check the status of a batch request. If the current instance of the Batch 
    'object was used to make the request, the batch_id is already known and is 
    'therefore optional.
     '@param string batch_id ID of the batch about which to enquire
     '@return  API get Request response string
    public function check_status(p_batch_id)
        if isNull(p_batch_id) or p_batch_id = "" then
            p_batch_id = batch_id
        end if

        call mc.getRequest("batches/" & p_batch_id, null, null)
        check_status = mc.getLastResponseBody()
    end function

    
     'Add an operation to the internal queue.
     '@param   string http_verb GET, POST, PUT, PATCH or DELETE
     '@param   string id ID for the operation within the batch
     '@param   string method URL of the API request method
     '@param   JSON arguments (usually your data)
    private sub queueOperation(http_verb, id, method, args)
        dim operation
        Set operation = New aspJSON
        with operation.data
            .add "operation_id", id
            .add "method", http_verb
            .add "path", method

            if not isNull( args ) then
                dim key
                if http_verb = "GET" then
                    key = "params"
                else
                    key = "body"
                end if
                
                .add key, Replace( args.JSONoutput(), vbCrLf, "")
            end if
        end with

        operations.data("operations").add ops, operation.data
        ops = ops + 1

    end sub

 end class
%>