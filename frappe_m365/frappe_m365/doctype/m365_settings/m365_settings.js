// Copyright (c) 2023, aptitudetech and contributors
// For license information, please see license.txt

frappe.ui.form.on('M365 Settings', {
	refresh: (frm) => {
		// Đặt giá trị mặc định cho m365_graph_url nếu chưa có
		if(!frm.doc.graph_api){
			frm.set_value('m365_graph_url', 'https://graph.microsoft.com/v1.0');
		}

		// Thiết lập bộ lọc cho trường role trong bảng module_settings
		frm.fields_dict['module_settings'].grid.get_field('role').get_query = function(doc, cdt, cdn) {
			let roles = [];
			$.each(cur_frm.doc.module_settings, function(index, row){
				roles.push(row.role);
			});

			return {
				filters:[
					['Role', 'name', 'not in', roles]
				]
			}
		}
	}
});

frappe.ui.form.on("M365 Groups Module Settings", {
	update_user: function(frm, cdt, cdn){
		// Hiển thị hộp thoại xác nhận trước khi cập nhật thành viên nhóm
		frappe.confirm(
			'<p>This action will update group membership on M365. Please confirm to proceed.</p>',
			function(){
				update_group_members(frm, cdt, cdn);
			},
			function(){}
		);
	}
});

function update_group_members(frm, cdt, cdn){
	let child = locals[cdt][cdn];
	// Kiểm tra nếu bản ghi chưa được lưu
	if(child.__islocal == 1){
		frappe.msgprint("Please save the form first.");
	}else if(child.role){
		// Gọi hàm Python để cập nhật thành viên nhóm
		frappe.call({
			"method": "frappe_m365.frappe_m365.doctype.m365_settings.m365_settings.update_group_members",
			"args": {"role": child.role, "group": child.default_group},
			"freeze": 1,
			"freeze_message": "<h4>Please wait while we are updating members in M365 Group...</h4>"
		})
	}else{
		frappe.msgprint("Please select a Role.");
	}
}