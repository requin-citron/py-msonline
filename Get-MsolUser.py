#!/usr/bin/env python
from azure.identity import InteractiveBrowserCredential, DefaultAzureCredential
import xml.etree.ElementTree as ET
import requests
import json
import uuid

def make_data(token):
    return '<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing"><s:Header><a:Action s:mustUnderstand="1">http://provisioning.microsoftonline.com/IProvisioningWebService/ListUsers</a:Action><a:MessageID>urn:uuid:'+str(uuid.uuid4())+'</a:MessageID><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><UserIdentityHeader xmlns="http://provisioning.microsoftonline.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><BearerToken xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService">Bearer '+token+'</BearerToken><LiveToken i:nil="true" xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService"/></UserIdentityHeader><User-Agent xmlns="http://becwebservice.microsoftonline.com/">msonline-powershell/msol</User-Agent><BecContext xmlns="http://becwebservice.microsoftonline.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><DataBlob xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService"></DataBlob><PartitionId xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService">326</PartitionId></BecContext><ClientVersionHeader xmlns="http://provisioning.microsoftonline.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><ClientId xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService">'+str(uuid.uuid4())+'</ClientId><Version xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService">1.2.183.81</Version></ClientVersionHeader><ContractVersionHeader xmlns="http://becwebservice.microsoftonline.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><BecVersion xmlns="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService">Version47</BecVersion></ContractVersionHeader><TrackingHeader xmlns="http://becwebservice.microsoftonline.com/">0b6658a8-82b5-4e0c-bd04-268e4bc336b2</TrackingHeader><a:To s:mustUnderstand="1">https://provisioningapi.microsoftonline.com/provisioningwebservice.svc</a:To></s:Header><s:Body><ListUsers xmlns="http://provisioning.microsoftonline.com/"><request xmlns:b="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration.WebService" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><b:BecVersion>Version16</b:BecVersion><b:TenantId i:nil="true"/><b:VerifiedDomain i:nil="true"/><b:UserSearchDefinition xmlns:c="http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration"><c:PageSize>500</c:PageSize><c:SearchString i:nil="true"/><c:SortDirection>Ascending</c:SortDirection><c:SortField>None</c:SortField><c:AccountSku i:nil="true"/><c:BlackberryUsersOnly i:nil="true"/><c:City i:nil="true"/><c:Country i:nil="true"/><c:Department i:nil="true"/><c:DomainName i:nil="true"/><c:EnabledFilter i:nil="true"/><c:HasErrorsOnly i:nil="true"/><c:IncludedProperties i:nil="true" xmlns:d="http://schemas.microsoft.com/2003/10/Serialization/Arrays"/><c:IndirectLicenseFilter i:nil="true"/><c:LicenseReconciliationNeededOnly i:nil="true"/><c:ReturnDeletedUsers i:nil="true"/><c:State i:nil="true"/><c:Synchronized i:nil="true"/><c:Title i:nil="true"/><c:UnlicensedUsersOnly i:nil="true"/><c:UsageLocation i:nil="true"/></b:UserSearchDefinition></request></ListUsers></s:Body></s:Envelope>'

def name(balise):
    out = str(balise).split("'")[1].split('}')
    if len(out) == 1:
        return str(balise).split("'")[1] 
    return out[1]
# Authenticate via Browser (Supports MFA)
#credential = InteractiveBrowserCredential()
credential = DefaultAzureCredential(exclude_interactive_browser_credential=False)

# Get Access Token for Microsoft Graph API
token = credential.get_token("https://graph.windows.net/.default").token

magie = make_data(token)

data = requests.post("https://provisioningapi.microsoftonline.com/provisioningwebservice.svc", data=magie, headers={"Content-Type":"application/soap+xml; charset=utf-8"})

namespaces = {
    's': 'http://www.w3.org/2003/05/soap-envelope',
    'c': 'http://schemas.datacontract.org/2004/07/Microsoft.Online.Administration',
    'd': 'http://schemas.microsoft.com/2003/10/Serialization/Arrays',
}

if "AccessDeniedException" in data.text:
    print("Access Denied")
    exit(1)

# Parse XML string
root = ET.fromstring(data.text)

out = list()

for user in root.findall(".//c:User",namespaces):
    user_dict = dict()
    for el in user:
        if el.text is None:
            tmp_lst = list()
            for subel in el:
                tmp_lst.append({name(subel):subel.text})
            user_dict[name(el)] = tmp_lst
        else:
            user_dict[name(el)] = el.text
    out.append(user_dict)

print(json.dumps(out, indent=4))