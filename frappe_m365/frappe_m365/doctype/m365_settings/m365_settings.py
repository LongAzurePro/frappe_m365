# Copyright (c) 2023, aptitudetech and contributors
# For license information, please see license.txt

import frappe
from frappe import _
from frappe.model.document import Document

class M365Settings(Document):
	pass

@frappe.whitelist()
def update_group_members(role, group):
	try:
		# Lấy danh sách người dùng có vai trò được chỉ định, ngoại trừ Administrator và Guest
		users = frappe.db.get_list(
			"Has Role",
			{
				"parenttype": "User", "role": role,
				"parent": ["not in", ["Administrator", "Guest"]]
			},
			"parent", ignore_permissions=1
		)

		# Lấy tài liệu nhóm M365
		group_doc = frappe.get_doc("M365 Groups", group)
		m365_group_members = [user.user for user in group_doc.m365_groups_member]
		
		# Thêm người dùng vào nhóm nếu họ chưa là thành viên
		for user in users:
			if user.parent not in m365_group_members:
				group_doc.append("m365_groups_member", {"user": user.parent})
	
		# Lưu tài liệu nhóm và cập nhật thành viên nhóm M365
		group_doc.save()
		group_doc.update_m365_groups_members()
	except Exception as e:
		frappe.msgprint(e)