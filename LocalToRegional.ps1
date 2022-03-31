#   Local to Regional API ticket generator script
#   written by Jax Sutton and Connor Pink
#define local and regional API tokens 
$Localtechnician_key = @{ 'authtoken' = 'key here' } #enter API keys for local and regional here.
$Regionaltechnician_key = @{ 'authtoken' = 'key here' } #enter API keys for local and regional here.
$localURL = "https://servicedesk.prhc.on.ca/" #Local system url here EX. https://domain.com
$regionalURL = "https://ourepic.ca/" #Regional system url here EX. https://domain.com
#temp location for attachments to be stored temporarily
$attachmentTempLocation = 'C:\Users\public\Documents\'

#Get all tickets with subcategory as NEeds Regional Ticket and send the IDs back
function GetTicketsIDs() {
    $url = $localURL+ "api/v3/requests"
    $input_data = @'
    {
        "list_info": {
            "row_count": 10,
            "start_index": 1,
            "sort_field": "subject",
            "sort_order": "asc",
            "get_total_count": false,
            "search_fields": {
                "subcategory.name": "Needs Regional Ticket"
            },
            "filter_by": {
                "name": "Open_System"
            }
            
        }
    }
'@
    $data = @{ 'input_data' = $input_data }
    $responses = Invoke-RestMethod -Uri $url -Method get -Body $data -Headers $Localtechnician_Key -ContentType "application/x-www-form-urlencoded"
    $IDs = @()
    foreach ($response in $responses.requests) {
        $IDs += $response.id
    }
    return $IDs
}

# Function to get info of local ticket
function GetTicket ($ticketNumber) {
    $url = $localURL+ "api/v3/requests/$ticketNumber"
    
    $response = Invoke-RestMethod -Uri $url -Method get -Headers $Localtechnician_Key
    return $response
    
}
#function to change local ticket to have regional ticket number link
function ChangeLocal($regionalID, $localTicketID) {
    #Powershell version - 5.1
    $url = $localURL+ "api/v3/requests/$localTicketID"
    $input_data = @"
    {
        'request': {
            'resolution': {
                'content': 'This ticket has been raised to the regional support group. The regional ticket submitted is $regionalID; please watch for communication from the regional support team, if you have any further questions please email navpilot@prhc.on.ca'
            },
            'udf_fields':  {
                'udf_long_6302': '$regionalID'
            },
            'status': {
                'name': 'Closed'
            }
        }
    }
"@
    $data = @{ 'input_data' = $input_data }
    $response = Invoke-RestMethod -Uri $url -Method put -Body $data -Headers $Localtechnician_Key -ContentType "application/x-www-form-urlencoded"
    return $response
}
#function to create the regional ticket with the correct info and return the regional ticket number
function CreateRegional ($info) {
    #values will need to be hard coded down below for program to work.
    $RegionalSupportGroup = $info.request.udf_fields.udf_pick_6903
    $subject = $info.request.subject
    $description = $info.request.description
    
    #If emails are cc'd on the ticket, put them in the emails to notify field on regional ticket
    if ($null -ne $info.request.email_cc){
        $emailCC = $info.request.email_cc
        #code to create the list of emails to notify from the cc list of the local ticket
        $newformattedCCs=@()
        for ($i = 0; $i -lt $emailCC.Length-1; $i++){
            $newformattedCCs += "'"+$emailCC[$i]+"',"
        }
        $newformattedCCs += "'"+$emailCC[$emailCC.Length-1]+"'"
    }
    
    #description manipulation to tend to the formatting from outlook on iphone
    $description = $description -replace '\n', ''
    $description = $description -replace '/div\\u003e', '/div\u003e\u003cbr /\u003e'

    $requester = $info.request.requester.email_id
    $urgency = $info.request.urgency.name 
    $impact = $info.request.impact.name
    
    # Some fields will have different values from system to system that mean the same thing so 
    # the values will be hard-coded here based on what they are in the API


    #Category & subcategory mapping for Local to Regional

        # Here is where the mapping of category, subcategory and group can be done for tickets. 
        #code here removed due to system security
    <#    
    if ($RegionalSupportGroup -eq "example group") {
        $category = "example category"
        #$subcategory = "example subcategory"
        $subcategoryID = "example id"
        $group = "example group" 
    }
    elseif ($RegionalSupportGroup -eq "different group") {
        $category = "different category"
        $subcategoryID = "different id"
        $group = "different group" 
    }
    ... and so on
     #>

    #Urgency mapping
    <#    EXAMPLE
    if ($urgency -eq "3 - Low") {
        $urgency = "Low - Minor Issue"
    }
    elseif ($urgency -eq "2 - Medium") {
        $urgency = "Medium - Performance Affected, No patient harm"
    }
    elseif ($urgency -eq "1 - High") {
        $urgency = "High - Work Stopped or System Down"
    }
 #>
    #Impact mapping
    <#  EXAMPLE
    if ($impact -eq "Low - Single User") {
        $impact = "Affecting Me"
    }
    elseif ($impact -eq "Medium - Single Unit") {
        $impact = "Affecting my Department"
    }
    elseif ($impact -eq "High - Multiple Units") {
        $impact = "Affecting Everyone"
    }
#>

    $url = $regionalURL+ "api/v3/requests"
    
    $input_data = @"
    {
        'request': {
            'subject': '$subject',
            'description': '$description',
            'requester': {
                'email_id': '$requester'
            },
            'status': {
                'name': 'Open'
            },
            'site': {
                'name' : 'Peterborough Regional Health Center'
            },
            'category': {
                'name': '$category'
            },
            'subcategory': {
                'id': '$subcategoryID'
            },
            'urgency': {
                'name': '$urgency'
            },
            'impact': {
                'name': '$impact'
            },
            'email_ids_to_notify':[ $newformattedCCs]
            ,
            'group': {
                'name': '$group'
            },
            'udf_fields':  {
                'udf_pick_601':  'E-Mail',
                'udf_sline_602':  '$requester',
                'udf_pick_304':  'Peterborough Health Center',
                'udf_pick_313':  'PRHC Information Technology',
                'udf_pick_301':  'Peterborough Regional Health Center'
            },
            'request_type':  {
                'name':  'Incident'
            }
        }
    }
"@
    $data = @{ 'input_data' = $input_data }
    $response = Invoke-RestMethod -Uri $url -Method post -Body $data -Headers $Regionaltechnician_Key -ContentType "application/x-www-form-urlencoded"
    return $response  
}
#if local ticket has attachment this function gets called
function HasSRC($ID, $regionalID, $info) {

    # code for if ticket has attachment 
    #-add note and change status to needs manual ticket generation
    #Powershell version - 5.1
   
    $url = $localURL+ "api/v3/requests/$ID/notes"        
    $input_data = @"
   
           {
               'note': {
                   'description': 'Regional Ticket With ID: $regionalID May Need Attachments To Be Manually Added.' ,
                   'mark_first_response': false,
                   'add_to_linked_requests': true,
                   'notify_technician': true
               }
           }
"@
    $data = @{ 'input_data' = $input_data }
    Invoke-RestMethod -Uri $url -Method post -Body $data -Headers $Localtechnician_Key -ContentType "application/x-www-form-urlencoded"
}

#function that gets and adds the attachment
function moveAttachments {
    param(
        [parameter(mandatory)]
        $PullTicketNumber,
        [parameter(mandatory)]
        $PutTicketNumber
    )

    #get all attachments in ticket
    $url = $localURL+ "api/v3/requests/$PullTicketNumber/attachments"

    $response = Invoke-RestMethod -Uri $url -Method get -ContentType "image/jpeg" -Headers $Localtechnician_Key

    #for each attachment on the ticket, launch the function to download them
    foreach ($image in $response.attachments) {
        downloadImage $image.id $PullTicketNumber $image.name
    }
}
#function to download the image by the attachment id on the ticket id
function downloadImage {
    param(
        [Parameter(Mandatory)]
        $imageID,
        [Parameter(Mandatory)]
        $ticketID,
        [parameter(mandatory)]
        $fileName
    )
    #get attachment
    $url = $localURL+ "api/v3/requests/$ticketID/attachments/$imageID/download"

    $fileLocation = $attachmentTempLocation + $fileName
    Invoke-RestMethod -Uri $url -Method get -ContentType "image/jpeg" -OutFile $fileLocation -Headers $Localtechnician_Key
    #upload the attachment to the regional
    uploadAttachment($fileLocation)
    #wait 2 seconds
    Start-Sleep -s 2
    #then delete the local file to keep computer clean.
    Remove-Item $fileLocation -Force

}
#function that uploads the attachment
function uploadAttachment() {

    param(

        [System.IO.FileInfo] $filename
    
    );
    #Input section starts 
    $server_url = $regionalURL
    #creating the formdata manually
    $boundary = [guid]::NewGuid().ToString()
    $filebody = [System.IO.File]::ReadAllBytes($filename)
    $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
    $filebodytemplate = $enc.GetString($filebody)
    [System.Text.StringBuilder]$contents = New-Object System.Text.StringBuilder
    $contents.AppendLine()
    $contents.AppendLine("--$boundary")
    $contents.AppendLine("Content-Disposition: form-data; name=""File""; filename=""$filename""")
    $contents.AppendLine()
    $contents.AppendLine($filebodytemplate)
    $contents.AppendLine("--$boundary--")
    $template = $contents.ToString()
    $inputData = "{`"attachment`": {`"request`": {`"id`": `"$PutTicketNumber`"}}}"
    $uri = $server_url+"api/v3/attachments?input_data=$inputData"
    
    $response = Invoke-RestMethod -Uri $uri -Method POST -Body $template -ContentType "multipart/form-data;boundary=$boundary" -Headers $Regionaltechnician_key
    $response
}



#function gets called if error occurs when regional ticket is created
function errorOccured($ID, $info, $regionalID) {

    $url = $localURL+ "api/v3/requests/$ID/notes"
        
    #add note ticket needs regional ticket to be manually created.
    $input_data = @"
    {
        'note': {
            'description': 'Ticket $regionalID needs Regional Ticket to be Manually Created. Automatic Regional Ticket Creation Has Failed',
            'mark_first_response': false,
            'add_to_linked_requests': true,
            'notify_technician': true
        }
    }
"@
    $data = @{ 'input_data' = $input_data }
    $response = Invoke-RestMethod -Uri $url -Method post -Body $data -Headers $Localtechnician_Key -ContentType "application/x-www-form-urlencoded"
    $response
    #add status manual regional ticket 
    $url = $localURL+ "api/v3/requests/$ID"
    $input_data = @"
{
    'request': {
        'status': {
            'name': 'Open'
        }
    }
}
"@
    $data = @{ 'input_data' = $input_data }
    $response = Invoke-RestMethod -Uri $url -Method put -Body $data -Headers $Localtechnician_Key -ContentType "application/x-www-form-urlencoded" 
}


#Call the function to get all the ticket IDs from local to work on
$localTicketIDs = GetTicketsIDs
Write-Host ("Working on Local Ticket Numbers: " + $localTicketIDs) #for debugging
foreach ($ID in $localTicketIDs) {
    #get the info of the local ticket
    $info = GetTicket($ID)
    try { 
        $newRegionalInfo = CreateRegional($info)

        if ($newRegionalInfo.request.description -like '*src*') {
            #add note that attachment needs to be added manually
            HasSRC $ID $newRegionalInfo.request.id $info
        }
        #change local ticket
        ChangeLocal $newRegionalInfo.request.id  $ID 
        #move attachments from local to regional function
        moveAttachments $ID $newRegionalInfo.request.id
    }
    #if any errors arise with creating regional ticket change local ticket to have not and status to represent such
    catch {  
        errorOccured $ID $info $newRegionalInfo.request.id
        Exit-PSHostProcess
    }
}
