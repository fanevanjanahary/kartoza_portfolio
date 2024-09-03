frappe.listview_settings['Portfolio'] = {
    onload: function(listview) {
        console.log("Portfolio list view loaded");

        listview.page.add_action_item(__('Export Portfolio'), function() {
            let selected = listview.get_checked_items();
            if (selected.length > 0) {
                show_export_dialog(selected);
            } else {
                frappe.msgprint(__('Please select at least one portfolio.'));
            }
        });
    }
};

function show_export_dialog(selected) {
    let d = new frappe.ui.Dialog({
        title: __('Choose Export Format'),
        fields: [
            {
                label: __('Format'),
                fieldname: 'format',
                fieldtype: 'Select',
                options: [
                    { label: 'PDF', value: 'pdf' },
                    { label: 'DOCX', value: 'docx' },
                    { label: 'HTML', value: 'html' },
                    { label: 'World Bank Format', value: 'world_bank' }
                ],
                default: 'pdf',
                reqd: 1
            }
        ],
        primary_action_label: __('Export'),
        primary_action: function(data) {
            d.hide();
            export_portfolios(selected, data.format);
        }
    });

    d.show();
}

function export_portfolios(selected, format) {
    console.log(format)
    frappe.call({
        method: 'portfolio.export.export_portfolio',
        args: {
            portfolio_names: JSON.stringify(selected.map(item => item.name)),
            format: format
        },
        callback: function(r) {
            if (r.message.status === 'success') {
                window.open(r.message.file_url, '_blank');
            } else {
                frappe.msgprint(__('Failed to export portfolio' + r.message.message));
            }
        }
    });
}