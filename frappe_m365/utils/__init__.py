from __future__ import absolute_import

import frappe
from frappe import _
import requests

#fetching oauth token
def get_oauth_token(settings):
    connected_app = frappe.get_doc("Connected App", settings.connected_app)
    oauth_token = connected_app.get_active_token(settings.connected_user)
    # print(oauth_token)
    # frappe.msgprint(str(oauth_token))
    # return None

    # tokens = frappe.get_all("Token Cache",filters={'user':settings.connected_user},fields=["user","access_token"])

    # for t in tokens:
    #     frappe.msgprint(str(t))

    if oauth_token == None:
        frappe.throw("Go to Chosen Connected App and Generate Token for Selected User On The Setting.")
    
    return oauth_token.get_password("access_token")

#making headers
def get_request_header(settings):
    access_token = get_oauth_token(settings)
    # frappe.msgprint(str(access_token))
    headers = {'Authorization': f'Bearer {access_token}'}
    return headers
    
#general api request
def make_request(request, url, headers, body=None):
    if(request == 'POST'):
        return requests.post(url, headers=headers, json=body)
    elif(request == 'PATCH'):
        return requests.patch(url, headers=headers, json=body)
    elif(request == 'GET'):
        return requests.get(url, headers=headers)
    elif(request == 'DELETE'):
        return requests.delete(url, headers=headers)
    elif(request == "PUT"):
        return requests.put(url, headers=headers, data=body)