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
				frm.trigger('create_team');
				frm.trigger('get_m365_members_on_server');
				frm.trigger('sync_office_365_links');
			}
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
	get_m365_members_on_server: function (frm) {
		// frm.add_custom_button(__("Get M365 members on Server"), function () {
		// 	if (frm.is_dirty()) {
		// 		frappe.msgprint("Please save the form first.")
		// 	} else {
		// 		frappe.call({
		// 			method: "get_m365_members_on_server",
		// 			freeze: 1,
		// 			freeze_message: "<h4>Please wait while we are retrieving members from M365</h4>",
		// 			doc: frm.doc,
		// 			callback: function (response) {
		// 				frappe.msgprint(JSON.stringify(response.message));
		// 			}
		// 		});
		// 	}
		// });

		add_user_to_m365 = (user_id) => {
			frappe.call({
				method: "add_user_to_m365",
				freeze: 1,
				freeze_message: "<h4>Please wait while we do the action.../h4>",
				doc: frm.doc,
				args:{
					user_id:user_id
				},
				callback: function (response) {
					console.log(response.message);
					frappe.msgprint(` ${JSON.stringify(response.message)} `);
					frm.reload_doc();
				}
			});
		}

		remove_member_from_m365 = (email) => {
			frappe.confirm("Are you sure to do this action? (This action cannot be turned back)",
				function (){
				frappe.call({
					method: "remove_member_from_m365",
					freeze: 1,
					freeze_message: "<h4>Please wait while we do the action.../h4>",
					doc: frm.doc,
					args:{
						email:email
					},
					callback: function (response) {
						console.log(response.message);
						frappe.msgprint(` ${JSON.stringify(response.message)} `);
						frm.reload_doc();
					}
					});
				}
			);
		}

		promote_member_to_m365_admin = (email) => {
			frappe.confirm("Are you sure to do this action? (This action cannot be turned back)",
				function (){
					frappe.call({
						method: "promote_member_to_m365_admin",
						freeze: 1,
						freeze_message: "<h4>Please wait while we do the action.../h4>",
						doc: frm.doc,
						args:{
							email:email
						},
						callback: function (response) {
							console.log(response.message);
							frappe.msgprint(` ${JSON.stringify(response.message)} `);
							frm.reload_doc();
						}
					});
				}
			)
			
		}

		remove_admin_from_m365 = (email) => {
			frappe.confirm("Are you sure to do this action? (This action cannot be turned back)",
				function (){
					frappe.call({
						method: "remove_admin_from_m365",
						freeze: 1,
						freeze_message: "<h4>Please wait while we do the action.../h4>",
						doc: frm.doc,
						args:{
							email:email
						},
						callback: function (response) {
							console.log(response.message);
							frappe.msgprint(` ${JSON.stringify(response.message)} `);
							frm.reload_doc();
						}
					});
				}
			)
			
		}

		add_member_form = () => {
			frappe.prompt([
                {'fieldname': 'email', 'fieldtype': 'Data', 'label': 'Nhập Email', 'reqd': 1}
            ],
            function(values){
                add_user_to_m365(values.email)
            },
            'Email',
            'Add member to M365');
		}

		frappe.call({
			method: "get_m365_members_on_server",
			freeze: 0,
			freeze_message: "<h4>Please wait while we are retrieving members from M365</h4>",
			doc: frm.doc,
			callback: function (response) {
				console.log(JSON.stringify(response.message));

				const data = response.message;

				let members_table_html = `
					<table class="table table-bordered">
						<thead>
							<tr>
								<th>Office 365 ID</th>
								<th>Name</th>
								<th>Designation</th>
								<th>Email</th>
								<th>Actions</th>
							</tr>
						</thead>
						<tbody>
				`;

				// Đổ dữ liệu vào bảng
				data.forEach(member => {
					members_table_html += `
						<tr>
							<td>${member.id}</td>
							<td>${member.displayName}</td>
							<td>${member.jobTitle ?? ""}</td>
							<td>${member.mail}</td>
							<td>
							<button onclick = "promote_member_to_m365_admin('${member.mail}')" class = "btn btn-primary">Promote to M365 Group Administrator</button>
							<br><br>
							<button onclick = "remove_member_from_m365('${member.mail}')" class = "btn btn-danger">Remove from M365 Group</button></td>
						</tr>
					`;
				});

				members_table_html += `
						</tbody>
					</table>
				`;
				members_table_html += `
					<button class="btn btn-primary" onclick = "add_member_form()">Add Member</button>
				`;

				// Đưa bảng HTML vào trường HTML Field
				frm.fields_dict['members_table_html'].$wrapper.html(members_table_html);
			}
		});

		frappe.call({
			method: "get_m365_admins_on_server",
			freeze: 0,
			freeze_message: "<h4>Please wait while we are retrieving members from M365</h4>",
			doc: frm.doc,
			callback: function (response) {
				console.log(JSON.stringify(response.message));

				const data = response.message;

				let admins_table_html = `
					<table class="table table-bordered">
						<thead>
							<tr>
								<th>Office 365 ID</th>
								<th>Name</th>
								<th>Designation</th>
								<th>Email</th>
								<th>Actions</th>
							</tr>
						</thead>
						<tbody>
				`;

				// Đổ dữ liệu vào bảng
				data.forEach(member => {
					admins_table_html += `
						<tr>
							<td>${member.id}</td>
							<td>${member.displayName}</td>
							<td>${member.jobTitle ?? ""}</td>
							<td>${member.mail}</td>
							<td>
							<button onclick = "remove_admin_from_m365('${member.mail}')" class = "btn btn-primary">Remove from M365 Group Administrator</button>
							<br><br>
							<button onclick = "remove_member_from_m365('${member.mail}')" class = "btn btn-danger">Remove from M365 Group</button></td>
						</tr>
					`;
				});

				admins_table_html += `
						</tbody>
					</table>
				`;

				// Đưa bảng HTML vào trường HTML Field
				frm.fields_dict['admins_table_html'].$wrapper.html(admins_table_html);
			}
		});

	},
	sync_office_365_links: function (frm){
		office_365_logo = `<img src="/assets/frappe/icons/social/office_365.svg" alt="Office 365">`
		teams_logo = `<img src="https://upload.wikimedia.org/wikipedia/commons/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg" alt="Microsoft Team" style = "width: 30px; height: 30px">`
		sharepoint_logo = `<img src="https://upload.wikimedia.org/wikipedia/commons/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg" alt="Sharepoint" style = "width: 30px; height: 30px">`
		outlook_logo = `<img src="https://upload.wikimedia.org/wikipedia/commons/d/df/Microsoft_Office_Outlook_%282018%E2%80%93present%29.svg" alt="Outlook" style = "width: 30px; height: 30px">`
		frm.add_custom_button(__(`Sync Office Links ${office_365_logo}`), function () {
			frappe.call({
				method: "sync_office_365_links",
				freeze: 0,
				freeze_message: "<h4>Please wait while we are create M365 Links...</h4>",
				doc: frm.doc,
				callback: function (r) {
					console.log(r.message);
	
					let teams_url = r.message.teams_url;
					let sharepoint_url = r.message.sharepoint_url
					let m365_url = r.message.m365_url
					frm.fields_dict.m365_team_redirect.$wrapper.html(`
						<button class = "btn btn-default" onclick="window.open('${teams_url}', '_blank')">Redirect to Teams ${teams_logo}</button>
					`);
					frm.fields_dict.m365_sharepoint_redirect.$wrapper.html(`
						<button class = "btn btn-default" onclick="window.open('${sharepoint_url}', '_blank')">Redirect to SharePoint ${sharepoint_logo}</button>
					`);
					frm.fields_dict.m365_group_redirect.$wrapper.html(`
						<button class = "btn btn-default" onclick="window.open('${m365_url}', '_blank')">Redirect to Outlook ${outlook_logo}</button>
					`);
				}
			});
		})

		if(frm.doc.m365_team_site !== null && frm.doc.m365_team_site !== undefined){
			frm.fields_dict.m365_team_redirect.$wrapper.html(`
				<button class = "btn btn-default" onclick="window.open('${frm.doc.m365_team_site}', '_blank')">Redirect to Teams ${teams_logo}</button>
			`);
		}
		if(frm.doc.m365_sharepoint_site !== null && frm.doc.m365_sharepoint_site !== undefined){
			frm.fields_dict.m365_sharepoint_redirect.$wrapper.html(`
				<button class = "btn btn-default" onclick="window.open('${frm.doc.m365_sharepoint_site}', '_blank')">Redirect to SharePoint ${sharepoint_logo}</button>
			`);
		}
		if(frm.doc.m365_group_site !== null && frm.doc.m365_group_site !== undefined){
			frm.fields_dict.m365_group_redirect.$wrapper.html(`
				<button class = "btn btn-default" onclick="window.open('${frm.doc.m365_group_site}', '_blank')">Redirect to Outlook ${outlook_logo}</button>
			`);
		}
	}
});
