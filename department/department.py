# Copyright (c) 2015, Frappe Technologies Pvt. Ltd. and Contributors
# License: GNU General Public License v3. See license.txt


import frappe
from frappe.utils.nestedset import NestedSet, get_root_of

from erpnext.utilities.transaction_base import delete_events

import urllib.parse


class Department(NestedSet):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from frappe.types import DF

		company: DF.Link
		department_name: DF.Data
		disabled: DF.Check
		is_group: DF.Check
		lft: DF.Int
		m365_group: DF.Link | None
		old_parent: DF.Data | None
		parent_department: DF.Link | None
		rgt: DF.Int
	# end: auto-generated types

	nsm_parent_field = "parent_department"

	def autoname(self):
		root = get_root_of("Department")
		if root and self.department_name != root:
			self.name = get_abbreviated_name(self.department_name, self.company)
		else:
			self.name = self.department_name

	def validate(self):
		if not self.parent_department:
			root = get_root_of("Department")
			if root:
				self.parent_department = root

	def before_rename(self, old, new, merge=False):
		# renaming consistency with abbreviation
		if frappe.get_cached_value("Company", self.company, "abbr") not in new:
			new = get_abbreviated_name(new, self.company)

		return new

	def on_update(self):
		if not (frappe.local.flags.ignore_update_nsm or frappe.flags.in_setup_wizard):
			super().on_update()

	def on_trash(self):
		super().on_trash()
		delete_events(self.doctype, self.name)
		
	@frappe.whitelist(methods=["GET"])
	def get_employees(self):
		employees = frappe.get_all(
			'Employee',
			filters={'department': self.name},
			fields=['name', 'employee_name', 'designation', 'department','user_id']
		)
		# frappe.msgprint(str(self))
		# frappe.msgprint(str(employees))
		return employees
	@frappe.whitelist()
	def get_m365_members_on_server(self):
		if self.m365_group:
			m365_group = frappe.get_doc("M365 Groups",self.m365_group)
			return m365_group.get_m365_members_on_server()
		else:
			return []
	
	@frappe.whitelist()
	def get_seperated_members(self):

		erpnext_members = self.get_employees()
		m365_members = self.get_m365_members_on_server()

		erpnext_emails = {member["user_id"] for member in erpnext_members}
		m365_lookup = {member["mail"]: member for member in m365_members}

		# Lọc ra các thành viên chỉ có trong ERPNext
		only_in_erpnext = [member for member in erpnext_members if member["user_id"] not in m365_lookup]

		# Lọc ra các thành viên chỉ có trong M365
		only_in_m365 = [member for member in m365_members if member["mail"] not in erpnext_emails]

		# both = [member for member in erpnext_members if member["user_id"] in m365_emails]

		both = [
			{
				**erp_member,  # Toàn bộ thuộc tính từ erpnext_members
				"office_365_id": m365_lookup[erp_member["user_id"]]["id"],  # Thuộc tính mới
				"office_365_name": m365_lookup[erp_member["user_id"]]["displayName"],  # Thuộc tính mới
			}
			for erp_member in erpnext_members
			if erp_member["user_id"] in m365_lookup  # Chỉ thêm nếu có trong m365_lookup
		]
		

		# frappe.msgprint(str({"only_in_erpnext":only_in_erpnext,"only_in_m365":only_in_m365}))

		return {"only_in_erpnext":only_in_erpnext,"only_in_m365":only_in_m365,"both":both}
	
	@frappe.whitelist()
	def add_erpnext_member_to_m365(self,user_id):
		if(self.m365_group):
			m365_group = frappe.get_doc("M365 Groups",self.m365_group)
			return m365_group.add_user_to_m365(user_id)
		else:
			frappe.throw("This Department doesn't have a M365 Group. Please create M365 Group for This Department.")
	
	@frappe.whitelist()
	def add_member_to_m365_via_power_automate(self,user_id):
		if(self.m365_group):
			m365_group = frappe.get_doc("M365 Groups",self.m365_group)
			return m365_group.add_member_to_m365_via_power_automate(user_id)
		else:
			frappe.throw("This Department doesn't have a M365 Group. Please create M365 Group for This Department.")
	
	@frappe.whitelist()
	def remove_member_from_m365(self,email):
		if(self.m365_group):
			m365_group = frappe.get_doc("M365 Groups",self.m365_group)
			return m365_group.remove_member_from_m365(email)
		else:
			frappe.throw("This Department doesn't have a M365 Group. Please create M365 Group for This Department.")

	@frappe.whitelist()
	def create_user_and_employee(self, email, full_name):
		# 1. Tạo User
		if not frappe.db.exists("User", email):
			user = frappe.get_doc({
				"doctype": "User",
				"email": email,
				"first_name": full_name,
				"enabled": 1,
				"send_welcome_email": 0  # Không gửi email chào mừng nếu không cần
			})
			user.insert(ignore_permissions=True)
			frappe.msgprint("User {0} created successfully.".format(email))
		else:
			frappe.msgprint("User {0} already exists.".format(email))

		employee = frappe.db.get_value("Employee", {"user_id": email}, "name")
		if employee:
			# Nếu Employee tồn tại, chuyển hướng đến giao diện chỉnh sửa
			return {
				"redirect_to": f"/app/employee/{employee}?full_name={full_name}&department={urllib.parse.quote(self.name)}"
			}

		# 2. Mở form Employee với các trường mặc định
		return {
			"redirect_to": f"/app/employee/new-employee-1?full_name={full_name}&department={urllib.parse.quote(self.name)}&user_id={email}",
		}



def on_doctype_update():
	frappe.db.add_index("Department", ["lft", "rgt"])


def get_abbreviated_name(name, company):
	abbr = frappe.get_cached_value("Company", company, "abbr")
	new_name = f"{name} - {abbr}"
	return new_name


@frappe.whitelist()
def get_children(doctype, parent=None, company=None, is_root=False):
	fields = ["name as value", "is_group as expandable"]
	filters = {}

	if company == parent:
		filters["name"] = get_root_of("Department")
	elif company:
		filters["parent_department"] = parent
		filters["company"] = company
	else:
		filters["parent_department"] = parent

	return frappe.get_all("Department", fields=fields, filters=filters, order_by="name")


@frappe.whitelist()
def add_node():
	from frappe.desk.treeview import make_tree_args

	args = frappe.form_dict
	args = make_tree_args(**args)

	if args.parent_department == args.company:
		args.parent_department = None

	frappe.get_doc(args).insert()

@frappe.whitelist()
def get_employees_by_department(department_name):
    employees = frappe.get_all(
        'Employee',
        filters={'department': department_name},
        fields=['name', 'employee_name', 'designation', 'department']
    )
    frappe.msgprint(department_name)
    frappe.msgprint(str(employees))
    return employees

@frappe.whitelist()
def unlink_employee_department(employee_name):
	# Fetch the Employee document
	employee = frappe.get_doc("Employee", employee_name)
	
	response = f"Unlinked employee {employee_name} from department {employee.department}"

	# Remove the department field
	employee.department = None

	# Save the changes
	employee.save()
	frappe.db.commit()

	return response