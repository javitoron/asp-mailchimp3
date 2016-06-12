<!--#include file="md5.asp"-->
<%
 ' Super-simple, minimum abstraction MailChimp API v3 wrapper
 ' MailChimp API v3: http://developer.mailchimp.com
 ' Based on PHP version: https://github.com/drewm/mailchimp-api
 '
 ' @author Javier Torón <javitoron@gmail.com>
 ' @version 0.1

class MailChimp

    private api_key
    private api_endpoint

    public verify_ssl
    private request_successful
    private last_error
    private last_response_headers
    private last_response_body
    private last_request
    private last_request_body

    Private Sub Class_Initialize()
        api_endpoint = "https://<dc>.api.mailchimp.com/3.0"
        verify_ssl = true
        request_successful = false
        last_error = ""
    End Sub

    Private Sub Class_Terminate()
    End Sub

    ' Create a new instance
    ' @param string api_key Your MailChimp API key
    public sub init(p_api_key)
        api_key = p_api_key

        if p_api_key = "" then
            Response.write "Empty MailChimp API key."
            response.end
        end if

        dim dashRevPos
        dashRevPos = instrrev(api_key, "-")

        if dashRevPos = 0 then
            Response.write "Invalid MailChimp API key supplied."
            response.end
        end if
        dashRevPos = len(api_key) - dashRevPos

        dim data_center
        data_center = right(api_key, dashRevPos )
        api_endpoint  = replace(api_endpoint, "<dc>", data_center)
    end sub


    'Create a new instance of a Batch request. Optionally with the ID of an existing batch.
    '@param string batch_id Optional ID of an existing batch, if you need to check its status for example.
    '@return Batch            New Batch object.
'    public function new_batch(batch_id = null)
'        return new Batch(this, batch_id)
'    end

     ' Convert an email address into a "subscriber hash" for identifying the subscriber in a method URL
     ' @param   string email The subscriber"s email address
     ' @return  string          Hashed version of the input
    public function subscriberHash(email)
        subscriberHash = MD5(lcase(email))
    end function

    'Was the last request successful?
    '@return bool  True for success, false for failure
    public function success()
        success = request_successful
    end function

    'Get the last error returned by either the network transport, or by the API.
    'If something didn"t work, this should contain the string describing the problem.
    '@return  array|false  describing the error
    public function getLastError()
        getLastError = last_error
    end function

    'Get an array containing the HTTP headers of the API response.
    '@return array  Assoc array with keys "headers" and "body"
    public function getLastResponseHeaders()
        getLastResponseHeaders = last_response_headers
    end function

    'Get an array containing the HTTP body of the API response.
    '@return array  Assoc array with keys "headers" and "body"
    public function getLastResponseBody()
        getLastResponseBody = last_response_body
    end function

    'Get an array containing the HTTP headers and the body of the API request.
    '@return array  Assoc array
    public function getLastRequest()
        getLastRequest = last_request
    end function

    public function getLastRequestBody()
        getLastRequestBody = last_request_body
    end function

    'Make an HTTP DELETE request - for deleting data
    '@param   string method URL of the API request method
    '@param   array args Assoc array of arguments (if any)
    '@param   int timeout Timeout limit for request in seconds
    '@return  array|false   Assoc array of API response, decoded from JSON
    public sub delete(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("delete", p_method, p_args, p_timeout)
    end sub

    'Make an HTTP GET request - for retrieving data
    '@param   string method URL of the API request method
    '@param   array args Assoc array of arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    '@return  array|false   Assoc array of API response, decoded from JSON
    public sub getRequest(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("get", p_method, p_args, p_timeout)
    end sub

    'Make an HTTP PATCH request - for performing partial updates
    '@param   string method URL of the API request method
    '@param   array args Assoc array of arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    '@return  array|false   Assoc array of API response, decoded from JSON
    public sub patch(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("patch", p_method, p_args, p_timeout)
    end sub
    
    'Make an HTTP POST request - for creating and updating items
    '@param   string method URL of the API request method
    '@param   array args Assoc array of arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    '@return  array|false   Assoc array of API response, decoded from JSON
    public sub post(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("post", p_method, p_args, p_timeout)
    end sub
    
    'Make an HTTP PUT request - for creating new items
    '@param   string method URL of the API request method
    '@param   array args Assoc array of arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    '@return  array|false   Assoc array of API response, decoded from JSON
    public sub putRequest(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("put", p_method, p_args, p_timeout)
    end sub
    
    'Performs the underlying HTTP request. Not very exciting.
    '@param  string http_verb The HTTP verb to use: get, post, put, patch, delete
    '@param  string method The API method to be called
    '@param  array args Assoc array of parameters to be passed
    '@param int timeout
    '@return array|false Assoc array of decoded result
    '@throws \Exception
    private sub makeRequest(http_verb, method, args, timeout)
        if IsNull(timeout) then
            timeout = 10
        end if

        dim url, req, params
        url = api_endpoint & "/" & method
        last_error = ""
        request_successful = false
        dim post_fields
        post_fields = ""

        last_request = "method => " & http_verb & ", path => " & method & _
            ", url => " & url & ", body => """", timeout => " & timeout

        set req = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        'Option SXH_OPTION_URL_CODEPAGE = 0
        req.setOption 0, 65001
        req.setTimeouts timeout * 500, timeout * 500, timeout * 1000, _
            timeout * 1000 'ms - resolve, connect, send, receive
        req.open http_verb, url & params, false ', "username", "password"
        req.setRequestHeader "Authorization", apiKey
        req.setRequestHeader "Content-Type", "application/json"

        if http_verb = "post" or http_verb = "patch" or http_verb = "put" then
                post_fields = args.JSONoutput()
        end if
            
        last_request_body = post_fields
        req.send post_fields

        select case http_verb
            case "get", "post", "put", "delete", "patch"
                'req.status = 400 --> ya existía el suscriptor
                if req.status >= 400 and req.status <= 599 then
                    request_successful = false
                    last_error = _
                        "Error: " & req.Status & " - " & req.statusText
                else
                    request_successful = true
                    last_error = ""
                end if

        end select

        last_response_headers = req.getAllResponseHeaders()
        last_response_body = req.responseText
    end sub
 end class
%>
