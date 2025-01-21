# Copyright (c) 2023, aptitudetech and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document
from frappe_m365.utils import get_request_header, make_request, convert_to_identifier, get_application_request_header
import time
import requests

M365 = "M365 Settings"
ContentType = {"Content-Type": "application/json"}


class M365Groups(Document):
    @frappe.whitelist()
    def run_m365_groups_flow(self):
        '''
            initiating the flow of M365 group, fetching the
            information of the groups, if exist in M365 then
            saving the info of group
            otherwise start new group creation
        '''
        if frappe.db.exists(M365):
            self._settings = frappe.get_single(M365)
            self.is_m365_group_exist()
        else:
            message = '''
				<p>
					Please update Oauth Settings in
					<a href='/app/m365-settings' style='color: #2490EF'>
						<b>M365 Settings</b>
					</a>
				</p>
			'''
            frappe.msgprint(message)

    def m365_groups_info(self):
        group_info = []
        headers = get_request_header(self._settings)
        url = f'{self._settings.m365_graph_url}/groups'
        groups = make_request('GET', url, headers, None)
        if groups.status_code == 200:
            for group in groups.json()['value']:
                if (group['displayName'] == self.m365_group_name):
                    group_info.append({'name': group['displayName'], 'id': group['id']})
        return group_info

    def get_user_info(self):
        '''
            getting user information which will be used
            for making owner of the group
        '''
        headers = get_request_header(self._settings)
        url = f'{self._settings.m365_graph_url}/me'
        res = make_request('GET', url, headers, None)#.json()
        if res.ok:
            return res.json()['id']
        else:
            frappe.throw(f"{res.json()['error']['message']}")

    def is_m365_group_exist(self):
        if not self.m365_group_id:
            group_info = self.m365_groups_info()
            if group_info:
                self.db_set('m365_group_id', group_info[0]['id'])
                self.m365_group_id = group_info[0]['id']
                frappe.msgprint('M365 Group has been successfully mapped.')
                self.initialize_M365_groups_services()
            else:
                self.create_m365_group()
        else:
            self.initialize_M365_groups_services()

    def create_m365_group(self):
        user_id = self.get_user_info()
        frappe.msgprint(f'{user_id}')
        template = self.template
        self.mailnickname = convert_to_identifier(self.name)
        frappe.msgprint(self.mailnickname)
        self.save()
        if template == "educationClass":
            # self._settings.connected_app = "73m034hajh"
            url = f'{self._settings.m365_graph_url}/education/classes'
            body = {
            "description": f"{self.m365_group_description}",
            "displayName": f"{self.m365_group_name}",
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "mailNickname": f"{self.mailnickname}",
            "securityEnabled": False,
                }
            headers = get_application_request_header(self._settings)
        else: 
            url = f'{self._settings.m365_graph_url}/groups'
            body = {
                "description": f"{self.m365_group_description}",
                "displayName": f"{self.m365_group_name}",
                "groupTypes": ["Unified"],
                "mailEnabled": True,
                "mailNickname": f"{self.mailnickname}",
                "securityEnabled": False,
                "owners@odata.bind": [f"{self._settings.m365_graph_url}/users/{user_id}"],
            }
            headers = get_request_header(self._settings)
        headers.update(ContentType)

        response = make_request('POST', url, headers, body)
        if (response.status_code == 201):
            self.db_set('m365_group_id', response.json()['id'])
            self.m365_group_id = response.json()['id']
            frappe.msgprint('M365 Group has been created successfully.')
            self.initialize_M365_groups_services()
        else:
            frappe.log_error("M365 Group Creation Error", response.text)
            frappe.msgprint(response.text)
        # self._settings.connected_app = "vs1tbfri1a"

        frappe.msgprint(f"Promoting {user_id} to group Administrator.")
        time.sleep(5)
        self.add_user_to_m365(user_id=user_id)
        self.promote_member_to_m365_admin(user_id=user_id)

        if template == "educationClass":
            url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}'
            headers = get_application_request_header(self._settings)
            headers.update(ContentType)
            body = {"owners@odata.bind": [f"{self._settings.m365_graph_url}/directoryObjects/{user_id}"]}
            response = make_request('PATCH', url, headers, body)
            self.create_team_for_m365_groups()
        
        

    def initialize_M365_groups_services(self):
        if not self.m365_sharepoint_id or not self.m365_sharepoint_site:
            # added sleep time so the group is properly initialize in MS365 for first itme
            time.sleep(10)

        msg = '''
				<p>The mapping of Frappe modules > M365 Group has started.
                You will be notified once the service is ready to use</p>
			'''
        frappe.msgprint(msg)
        self.create_sharepoint_service()

    def create_sharepoint_service(self):
        if not self.m365_sharepoint_site:
            headers = get_request_header(self._settings)
            url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/sites/root'

            response = make_request('GET', url, headers, None)
            if (response.status_code == 200 or response.ok):
                self.db_set("m365_sharepoint_site", response.json()['webUrl'])
                self.m365_sharepoint_site = response.json()['webUrl']
            else:
                frappe.log_error("sharepoint webUrl error", response.text)
        self.map_sharepoint_id()

    def map_sharepoint_id(self):
        if not self.m365_sharepoint_id:
            headers = get_request_header(self._settings)
            url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/drive/root/children'

            response = make_request('GET', url, headers, None)
            if (response.ok):
                for items in response.json()['value']:
                    if self.name == items["name"]:
                        self.db_set("m365_sharepoint_id", items['id'])
                        self.m365_sharepoint_id = items['id']

            if not self.m365_sharepoint_id:
                headers = get_request_header(self._settings)
                body = {
                    "name": f'{self.name}',
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "rename"
                }
                url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/drive/items'
                response = make_request('POST', url, headers, body)

                if (response.ok):
                    self.db_set("m365_sharepoint_id", response.json()["id"])
                    self.m365_sharepoint_id = response.json()["id"]
                else:
                    frappe.log_error("sharepoint id error", response.text)

        if self.m365_sharepoint_site and self.m365_sharepoint_id:
            frappe.enqueue("frappe_m365.utils.sharepoint.trigger_sharepoint",
                           queue="long", request="MAP", group=self, timeout=-1)
            
    @frappe.whitelist()
    def update_m365_groups_members(self):
        if (self.m365_group_id):
            self._settings = frappe.get_single(M365)
            self.add_members_in_group()
            self.delete_members_in_group()
        elif (not self.group_idm365_group_id):
            frappe.msgprint("Please <b>Connect to M365 Groups</b> first.")

    def get_group_member_list(self):
        """
            fetching M365 Group member list
        """
        members = []
        url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/members'
        headers = get_request_header(self._settings)
        response = make_request('GET', url, headers, None)
        if (response.ok):
            for member in response.json()['value']:
                members.append({"mail": member['mail'], "id": member['id']})
        else:
            frappe.log_error("M365 Group member fetching error", response.text)
        return members

    def get_m365_users_list(self):
        """
            fetching user list with id from organization
        """
        users = []
        url = f'{self._settings.m365_graph_url}/users'
        headers = get_request_header(self._settings)
        headers.update(ContentType)

        response = make_request('GET', url, headers, None)
        if (response.ok):
            for user in response.json()['value']:
                users.append({"mail": user['mail'], "id": user['id']})
        else:
            frappe.log_error("M365 users fetching error", response.text)
        return users

    def add_members_in_group(self):
        """
            checking the listed member(s) exists in the
            organization then adding listed member(s) in
            group if not preset in group member list
        """
        org_users = self.get_m365_users_list()
        org_users_mails = [user['mail'] for user in org_users]
        group_members = self.get_group_member_list()
        group_members_mails = [member['mail'] for member in group_members]

        members_not_in_group = []
        members_not_in_org = []
        for member in self.m365_groups_member:
            mail = member.user
            if mail not in group_members_mails and mail in org_users_mails:
                member_id = ''.join([user['id'] for user in org_users if user['mail'] == mail])
                members_not_in_group.append(f'{self._settings.m365_graph_url}/directoryObjects/{member_id}')
            elif (mail not in org_users_mails):
                members_not_in_org.append(mail)

        if members_not_in_group:
            url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}'
            headers = get_request_header(self._settings)
            headers.update(ContentType)
            body = {"members@odata.bind": members_not_in_group}

            response = make_request('PATCH', url, headers, body)
            if not response.ok:
                frappe.log_error("M365 Group member(s) update Error", response.text)
                frappe.msgprint(response.text)
            else:
                frappe.msgprint("Group Member(s) has been updated successfully")
        else:
            frappe.msgprint("Group Member(s) list is up-to date")

        if members_not_in_org:
            msg = """
				<p>At least {0} user(s) is not part of your organization:</p>
				<p><b>{1}</b></p>
				<p>User(s) in this list will not be added to the M365 Group.</p>
			""".format(len(members_not_in_org), "<br>".join(members_not_in_org))
            frappe.msgprint(msg)

    def delete_members_in_group(self):
        delete_member_from_group = []
        group_members = self.get_group_member_list()
        member_data = [member.user for member in self.m365_groups_member]
        for member in group_members:
            if (member['mail'] not in member_data):
                delete_member_from_group.append(member['id'])

        for member in delete_member_from_group:
            url = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/members/{member}/$ref'
            headers = get_request_header(self._settings)

            response = make_request('DELETE', url, headers, None)
            if not response.ok:
                frappe.log_error("M365 Group member(s) delete Error", response.text)
    
    # Add Team for M365 Group
    @frappe.whitelist()
    def create_team_for_m365_groups(self):
        self._settings = frappe.get_single(M365)

        group_id = self.m365_group_id  # Đây là ID của nhóm M365 đã tạo trước đó

        url = f'{self._settings.m365_graph_url}/groups/{group_id}/team'

        headers = get_request_header(self._settings)
        headers.update(ContentType)


        check_response = requests.get(url, headers=headers)
        check_response_data = check_response.json()
        if(check_response_data.get("id")):
            frappe.msgprint(f'Team for this M365 Group exists(ID: {check_response_data["id"]})')
            self.m365_team_id = check_response_data["id"]
            self.save()
            frappe.db.commit()

        else:
            url = f'{self._settings.m365_graph_url}/teams'
            body = {
                "template@odata.bind": f"{self._settings.m365_graph_url}/teamsTemplates('{self.template}')",
                "group@odata.bind": f"{self._settings.m365_graph_url}/groups('{group_id}')",
                "memberSettings": {
                    "allowCreateUpdateChannels": True,
                    "allowDeleteChannels": True
                },
                "messagingSettings": {
                    "allowUserEditMessages": True,
                    "allowUserDeleteMessages": True
                },
                "funSettings": {
                    "allowGiphy": True,
                    "allowStickersAndEmojis": True
                },
            }
            
            if self.template == "educationClass":
                response = requests.post(url,headers=headers,json=body)
            else:
                response = requests.post(url,headers=headers,json=body)

            # response = make_request("PUT" if  else "POST", url, headers, body)

            if response.status_code == 201:
                frappe.msgprint(response.text)

                response_data = response.json()
                self.m365_team_id = response_data["id"]
                self.save()
                frappe.db.commit()
                
                frappe.msgprint("Microsoft Teams group has been created successfully.")
                return "Microsoft Teams group has been created successfully."
            else:
                frappe.log_error("Teams Group Creation Error", response.text)
                frappe.msgprint(response.text)
    @frappe.whitelist()
    def get_m365_members_on_server(self):
        self._settings = frappe.get_single(M365)

        group_id = self.m365_group_id  # Đây là ID của nhóm M365 đã tạo trước đó

        url = f'{self._settings.m365_graph_url}/groups/{group_id}/members'

        headers = get_request_header(self._settings)
        headers.update(ContentType)

        response = requests.get(url, headers=headers)
        response_dict = response.json()
        
        if response.status_code == 200 and response_dict.get('value'):
            # frappe.msgprint(str(response_dict["value"]))
            return response_dict["value"]
            
        else:
            return []
            # frappe.throw(f'Error: {response.text}')
    
    @frappe.whitelist()
    def add_user_to_m365(self,email=None,user_id=None):
        # email = frappe.get_value("User", {"email": user_id}, "email")
        # URL endpoint để thêm thành viên vào nhóm
        url = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/members/$ref"

        headers = get_request_header(self._settings)
        headers.update(ContentType)


        if not user_id:
            self._settings = frappe.get_single(M365)

            # Lấy user ID từ email
            user_url = f"https://graph.microsoft.com/v1.0/users/{email}"

            
            user_response = requests.get(
                user_url,
                headers=headers
            )
        
            if user_response.status_code != 200:
                # frappe.response["http_status_code"] = 400
                
                return str({"error": f"Cannot find user with email {email}", "details": user_response.json()})
                # return f"Cannot find user with email {email}"

            # office_user_id = user_response.json().get("id")
            user_id = user_response.json().get("id")
            frappe.msgprint('')
        
        office_user_id = user_id
        # frappe.msgprint(user_response.text)

        if not office_user_id:
            # frappe.response["http_status_code"] = 400
            return str({"error": "User ID not found in the response"})
        
            # return f"Cannot find user with email {email}"

        # Dữ liệu để thêm người dùng vào nhóm
        payload = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{office_user_id}"
        }

        # Gửi yêu cầu POST để thêm thành viên vào nhóm
        response = requests.post(
            url,
            headers=headers,
            json=payload
        )

        # Xử lý phản hồi
        if response.status_code == 204:

            # frappe.response["http_status_code"] = 200

            # After that email added to M365, create a new M365 Group Member for M365 Group Doc
            if frappe.db.exists("User", {"email": email}):
                self.append("m365_groups_member",{
                    "doctype":"M365 Groups Member",
                    "user":email,
                })

                self.save()
                frappe.db.commit()

            return {"success": f"User {email} added to group {self.name}"}
            # return f"User {email} added to group {self.name}"
        else:
            # frappe.response["http_status_code"] = 400

            return {"error": "Failed to add user to group", "details": response.json()}

            # return f"Failed to add user to group {response.text}"

    @frappe.whitelist()
    def add_member_to_m365_via_power_automate(self,user_id):

        #Is this an innovation??? I'd think so
    
        email = frappe.get_value("User", {"email": user_id}, "email")

        settings = frappe.get_single(M365)

        connected_app = frappe.get_doc("Connected App", settings.connected_power_automate)
        oauth_token = connected_app.get_active_token(settings.connected_user)

        headers = {'Authorization': f'Bearer {oauth_token.get_password("access_token")}'}
        # headers.update(ContentType)

        # URL của trigger từ Power Automate
        url = "https://prod-36.southeastasia.logic.azure.com:443/workflows/a33f2efc32854362994f3fb1e65d6093/triggers/manual/paths/invoke?api-version=2016-06-01"

        # Thông tin người dùng cần thêm
        data = {
            "email": email,
            "group_id": self.m365_group_id
        }

        # Gọi HTTP request đến Power Automate
        response = requests.post(url, json=data, headers=headers)

        # Kiểm tra kết quả
        if response.status_code == 200:
            return "User added to the group successfully!"
        else:
            return f"Failed to add user {response.text}"

        
    @frappe.whitelist()
    def remove_member_from_m365(self,email):
        self._settings = frappe.get_single(M365)

        # 1. Lấy member-id từ email
        headers = get_request_header(self._settings)
        headers.update(ContentType)

        # 1. Tìm người dùng bằng email
        user_url = f'https://graph.microsoft.com/v1.0/users/{email}'
        response = requests.get(user_url, headers=headers)

        if response.status_code == 200:
            user_data = response.json()
            user_id = user_data['id']
            
        else:
            return f'Error when retrieving user infomation ({response.status_code}): {response.json()}'

        # 2. Loại bỏ người dùng khỏi nhóm
        remove_url = f'https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/members/{user_id}/$ref'
        remove_response = requests.delete(remove_url, headers=headers)

        if remove_response.status_code == 204:

            for m in self.m365_groups_member:
                if m.user == email:
                    self.m365_groups_member.remove(m)
                    self.save()
                    frappe.db.commit()
                    break

            return f'User {email} is successfully removed from Group {self.m365_group_id}'
        else:
            return f'Error when removing user from this group: {remove_response.status_code} {remove_response.json()}'
    
    @frappe.whitelist()
    def get_m365_admins_on_server(self):
        self._settings = frappe.get_single(M365)

        group_id = self.m365_group_id  # Đây là ID của nhóm M365 đã tạo trước đó

        url = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/owners"

        headers = get_request_header(self._settings)
        headers.update(ContentType)

        response = requests.get(url, headers=headers)
        response_dict = response.json()
        
        if response.status_code == 200 and response_dict.get('value'):
            # frappe.msgprint(str(response_dict["value"]))
            return response_dict["value"]
            
        else:
            return []
        
    @frappe.whitelist()
    def promote_member_to_m365_admin(self,email=None,user_id=None):

        # Headers cho request

        self._settings = frappe.get_single(M365)

        headers = get_request_header(self._settings)

        if not user_id:
            user_lookup_url = f"https://graph.microsoft.com/v1.0/users/{email}"
            

            # 1. Tìm ID của user từ email
            response = requests.get(user_lookup_url, headers=headers)
            if response.status_code == 200:
                user_data = response.json()
                user_id = user_data.get("id")
                # frappe.msgprint(f"User ID: {user_id}")
            else:
                return f"Email {email} not found in this tenant!"

        # # 2. Thêm user vào nhóm nếu chưa có
        # add_member_url = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/members/$ref"
        member_data = {
            "@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"
        }

        # response = requests.post(add_member_url, headers=headers, json=member_data)
        # if response.status_code in [200, 204]:  # 204: No Content
        #     frappe.msgprint("Email {email}")
        # elif response.status_code == 400 and "already exists" in response.text:
        #     frappe.msgprint("Người dùng đã có trong nhóm.")
        # else:
        #     frappe.msgprint(f"Lỗi khi thêm thành viên: {response.text}")

        # 3. Thêm user vào danh sách owner (quản trị viên)
        add_owner_url = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/owners/$ref"
        response = requests.post(add_owner_url, headers=headers, json=member_data)

        if response.status_code in [200, 204]:
            return f"Email {email} became Administrator of Group {self.m365_group_id}!"
        else:
            return f"Error when asigning Administrator: {response.text}"
    
    @frappe.whitelist()
    def remove_admin_from_m365(self,email):
        user_lookup_url = f"https://graph.microsoft.com/v1.0/users/{email}"
        # Headers cho request

        self._settings = frappe.get_single(M365)

        headers = get_request_header(self._settings)

        # 1. Tìm ID của user từ email
        response = requests.get(user_lookup_url, headers=headers)
        if response.status_code == 200:
            user_data = response.json()
            user_id = user_data.get("id")
            # frappe.msgprint(f"User ID: {user_id}")
        else:
            return f"Email {email} not found in this tenant!"

        # 2. Xóa tài khoản khỏi danh sách admin (quản trị viên)
        remove_owner_url = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}/owners/{user_id}/$ref"

        response = requests.delete(remove_owner_url, headers=headers)

        if response.status_code in [200, 204]:
            return f"Email {email} removed from Administrators list of Group {self.m365_group_id}!"
        else:
            return f"Error when asigning Administrator: {response.text}"
        
    

    @frappe.whitelist()
    def sync_office_365_links(self):
        self._settings = frappe.get_single(M365)

        # 1. Lấy member-id từ email
        headers = get_request_header(self._settings)
        headers.update(ContentType)

        result = {}

        if self.m365_team_id:
            team_api = f"https://graph.microsoft.com/v1.0/teams/{self.m365_team_id}"
            response = requests.get(team_api,headers=headers)
            response = response.json()

            self.m365_team_site = response.get("webUrl")
            self.save()

            result["teams_url"] = response.get("webUrl")
        if self.m365_sharepoint_id:
            sharepoint_api = f'{self._settings.m365_graph_url}/groups/{self.m365_group_id}/sites/root'
            response = requests.get(sharepoint_api,headers=headers)
            response = response.json()

            self.m365_sharepoint_name = response.get("webUrl")
            self.save()

            result["sharepoint_url"] = response.get("webUrl")
        if self.m365_group_id and self.template == "standard":
            m365_api = f"https://graph.microsoft.com/v1.0/groups/{self.m365_group_id}"
            response = requests.get(m365_api,headers=headers)

            frappe.msgprint(response.text)
            response = response.json()
            mailnickname,company_site = response["mail"].split("@")
            m365_url = f"https://outlook.office.com/groups/{company_site}/{mailnickname}/mail"

            self.m365_group_site = m365_url
            self.save()

            result["m365_url"] = m365_url

        return result

    @frappe.whitelist()
    def get_teams_templates(self):
        self._settings = frappe.get_single(M365)

        # 1. Lấy member-id từ email
        headers = get_request_header(self._settings)
        headers.update(ContentType)

        url = "https://graph.microsoft.com/beta/teamwork/teamTemplates"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.text
        else:
            return f"Error: {response.status_code}, {response.text}"




@frappe.whitelist()
def create_m365_group_for_any_doc(doc,members_doctype=None,members_search_field=None,*args,**kwargs):
    """

    """
    doc = eval(doc)
    m365_members = []
    
    # Tạo doc M365
    group_doc = frappe.get_doc("M365 Groups",f"{doc['name']}" + f" - {doc['company']}" if doc.get('company') else "") if frappe.db.exists("M365 Groups", f"{doc['name']}" + f" - {doc['company']}" if doc.get('company') else "") else frappe.get_doc({
            "doctype":"M365 Groups",
            "m365_group_name":f"{doc['name']}" + f" - {doc['company']}" if doc.get('company') else "",
            "m365_group_description":f"M365 Group for " + f"{doc['name']}" + f" - {doc['company']}" if doc.get('company') else "",
            "enable":True
    })
    group_doc.save()
    frappe.db.commit()

    if frappe.db.exists("DocField", {"parent": doc["doctype"], "fieldname": "m365_group"}):
        doc_obj = frappe.get_doc(doc["doctype"],doc["name"])
        doc_obj.m365_group = group_doc
        doc_obj.save()
        frappe.db.commit()
    
    #Chạy luồng tạo nhóm M365 trên Graph Microsoft
    group_doc.run_m365_groups_flow()
    group_doc.save()
    frappe.db.commit()

    #Đồng bộ thành viên
    if (members_doctype and members_search_field):
        members = frappe.get_all(
            members_doctype,
            filters={members_search_field: doc['name']},
            fields=['user_id']
        )
        # frappe.msgprint(str(members))
        # [frappe.msgprint(str(item)) for item in members]
        # members = [frappe.get_doc({"doctype": "M365 Groups Member","user": item["user_id"]}) for item in members]
        for m in members:
            if not frappe.db.exists("M365 Groups Member",m["user_id"]):
                new_mem = frappe.get_doc({"doctype":"M365 Groups Member","user":m["user_id"]})
                # group_doc.m365_groups_member.append(""new_mem)
                group_doc.append("m365_groups_member",{
                    "doctype":"M365 Groups Member",
                    "user":m["user_id"],
                })

                # new_mem.save()
                group_doc.save()
                frappe.db.commit()
        # frappe.msgprint(str(group_doc.__dict__))
        for m in group_doc.m365_groups_member:
            frappe.msgprint(str(m.__dict__))
    

    frappe.msgprint(f"M365 Group created for {doc['name']}")
    pass