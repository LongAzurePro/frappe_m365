// Copyright (c) 2023, aptitudetech and contributors
// For license information, please see license.txt

frappe.ui.form.on('M365 Groups', {
	refresh: function(frm) {
		frm.trigger('set_default_values');
		if (frm.doc.enable && !frm.doc.__islocal) {
			if(!frm.doc.m365_group_id){
				frm.trigger('add_connect_button');
			}else{
				frm.trigger('update_group_members');
				if(frm.doc.m365_team_id){
					frm.trigger('open_msteam');
				}else{
					frm.trigger('create_team');	
				}			
				frm.trigger('handle_add_employee');
				frm.trigger('test');
			}
		}
	},
	handle_add_employee: function(frm){
		if(frm.doc.department_id && frm.doc.company){
			frappe.call({
				method: "frappe.client.get_list",
				args :{
				"doctype": "Employee",
				"fields": ["name","user_id","department","company"],
				"filters": [
					["department", "=", frm.doc.department_id],
					["company", "=", frm.doc.company]
				]
				},
				callback: function(r){
					let userIds = r.message.map(emp => emp.user_id).filter(user_id => user_id);
					frappe.call({
						method: "frappe.client.get_list",
						args:{
							doctype: "User",
							filters: {
								"name": ["in", userIds]
							},
							fields: ["name"]
						},
						callback: function(r){
							if(r.message){
								if (!frm.doc["m365_groups_member"] || !frm.doc["m365_groups_member"].length > 0){
								frm.clear_table("m365_groups_member");
								r.message.forEach(user=>{
									let row =frm.add_child("m365_groups_member");
									row.user = user.name;
								});
								frm.refresh_field("m365_groups_member");
							}
						}
						}
					})
				}
			})
		}
	},
	set_default_values: function (frm) {
		if (!frm.doc.m365_group_description) {
			frm.set_value('m365_group_description', 'This group has been created from Frappe-M365');
		}

		if (frm.doc.group_id) {
			frm.toggle_enable("mailnickname", 1);
		}
	},

	m365_group_name: function (frm) {
		let parts = frm.doc.m365_group_name.split(" ");
		let abbr = $.map(parts, function (p) {
			return p ? p.substr(0, 1) : null;
		}).join("");
		frm.set_value("mailnickname", abbr.toLowerCase());
	},

	add_connect_button: function (frm) {
		frm.add_custom_button(__("Connect to M365 Groups"), function () {
			if (!frm.is_dirty()) {
				frappe.call({
					method: "run_m365_groups_flow",
					freeze: 1,
					freeze_message: "<h4>Please wait while we are connecting and mapping with M365 groups...</h4>",
					doc: frm.doc,
					callback: function (r) {
						frm.set_df_property("template","read_only",1);
						frm.refresh_field("template");
						frm.reload_doc();
					}
				});
			} else {
				frappe.msgprint("Please save the form first.")
			}
		});
	},

	update_group_members: function (frm) {
		frm.add_custom_button(__("Update M365 Group Member(s)"), function () {
			if (frm.is_dirty()) {
				frappe.msgprint("Please save the form first.")
			} else {
				frappe.call({
					method: "update_m365_groups_members",
					freeze: 1,
					freeze_message: "<h4>Please wait while we are updating members in M365 Group...</h4>",
					doc: frm.doc,
					callback: function (r) {
						frm.reload_doc();
					}
				});
			}
		});
	},
	create_team: function (frm) {
		frm.add_custom_button(__("Create Team for M365 Group"), function () {
			if (frm.is_dirty()) {
				frappe.msgprint("Please save the form first.")
			} else {
				frappe.call({
					method: "create_team_for_m365_groups",
					freeze: 1,
					freeze_message: "<h4>Please wait while we are creating a Team for M365 Group...</h4>",
					doc: frm.doc,
					callback: function (r) {
						frm.reload_doc();
					}
				});
			}
		});
	},
	open_msteam: function(frm){
		frm.add_custom_button(__("Open Microsoft Teams"), function(){
		frappe.call({
			method: "open_msteam",
			doc: frm.doc,
			freeze: 1,
			freeze_message: "<h4>Please wait while we are opening Microsoft Teams...</h4>",
			callback: function(r){
				if(r.message){
					frappe.msgprint(`${r.message}`)
					window.open(r.message, "_blank");
				}
			}
		})
	});
	},

	test: function(frm){
		frappe.call({
			method: "m365_groups_info",
			doc: frm.doc,
			freeze: 1,
			freeze_message: "<h4>Please wait while we are testing...</h4>",
			callback: function(r){
				if(r.message){
					if(r.message.length === 0){
						frm.set_value("m365_group_id", "");
						frm.set_value("m365_team_id", "");
						frm.save();
						frm.refresh_field("m365_group_id");
						frm.refresh_field("m365_team_id");
					}
				}
			}
		})
	}
});
