import frappe
from frappe.utils.pdf import get_pdf
from frappe import _
import io
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from bs4 import BeautifulSoup
from docx.shared import Inches, Pt
from datetime import datetime
import tempfile
import os

@frappe.whitelist()
def export_portfolio(portfolio_names, format):
    if not portfolio_names:
        frappe.throw(_("No portfolio names provided"))

    file_data_list = []
    content = generate_html_content(portfolio_names)

    if format == "pdf":
        file_data = get_pdf(content)
        file_extension = "pdf"
    elif format == "docx":
        file_data = generate_docx(content)
        file_extension = "docx"
    elif format == "world_bank":
        file_data = worldbank_format(portfolio_names)
        file_extension = "docx"
    elif format == "html":
        file_data = generate_html_file(content)
        file_extension = "html"
    else:
        frappe.throw(_("Unsupported file format"))

    file_data_list.append(file_data)

    # Combine file_data_list into a single file if needed, for simplicity, return the first file.
    combined_file_data = file_data_list[0]

    # Generate a default file name with timestamp
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    default_file_name = f"portfolio_export_{timestamp}.{file_extension}"

    # Create a new File document to save the generated file
    file_doc = frappe.get_doc({
        "doctype": "File",
        "file_name": default_file_name,
        "is_private": 1,
        "content": combined_file_data
    })
    file_doc.insert()

    return {
        "status": "success",
        "message": f"Portfolios exported successfully.",
        "file_url": file_doc.file_url
    }


def generate_html_content(portfolios):
    project_details = ""
    portfolio_names = frappe.parse_json(portfolios)
    for docname in portfolio_names:
        portfolio = frappe.get_doc("Portfolio", docname)
        technologies_list = ""
        for tech in portfolio.technologies:
            technologies_list += f"<li>{tech.technology}</li>"
			
        client_contact = ""
        if portfolio.contact != "":
            client_contact = portfolio.contact
        else:
            client_contact = "Unavailable"

        client_reference = ""
        if portfolio.client_reference != "":
            client_reference = portfolio.client_reference
        else:
            client_reference = "Unavailable"
		
        client_logo = ""
        if portfolio.client_logo != "":
            client_logo =frappe.utils.get_url() + portfolio.client_logo

			
        services_list = ""
        for service in portfolio.services_listed:
            services_list += f"<li>{service.service}</li>"

        absolute_url = frappe.utils.get_url() 

        time = absolute_url + "/assets/portfolio/images/time.png"
        location = absolute_url + "/assets/portfolio/images/location.png"
        person = absolute_url + "/assets/portfolio/images/person.png"
        footer = absolute_url + "/assets/portfolio/images/footer.png"

        images_list = ""
        for image in portfolio.images:
            if image and image.website_image:
                image_url = image.website_image
                # Check if the URL is missing a schema and prepend one if necessary
                if not image_url.startswith(('http://', 'https://')):
                    image_url = frappe.utils.get_url() + image_url
                images_list += f'<img src="{image_url}" alt="Screenshot" style="width:100%;height:100%;object-fit:contain;padding:10px"><br>'

        project_details += f"""
		<h3 style="color:#f4b340;text-align:center">Kartoza Project Sheet</h3>
        <h2 style="text-align:center">{portfolio.title}</h2>
		<div>
            <hr style=" border: 8px solid #f4b340; width: 90px; margin:auto !important;">
		</div>
		<br><br>
		<div style="display: flex; width: 100%;">
            <div style="flex: 1; margin: 0; text-align:center; border: 1px solid gray; padding: 10px; display: flex; flex-direction: column; justify-content: center;">
                <div style="text-align:center">
                    <img src="{person}" alt="Project Image" style="width:80px;height:auto;">
				</div>
                <p>Client: {portfolio.client}</p>
            </div>
            <div style="flex: 1; margin: 0; text-align:center; border: 1px solid gray; padding: 10px; display: flex; flex-direction: column; justify-content: center;">
                <div >
                    <img src="{location}" alt="Project Image" style="width:80px;height:auto;text-align:center">
				</div>
                <p>Location: {portfolio.location}</p>
            </div>
            <div style="flex: 1; margin: 0; text-align:center; border: 1px solid gray; padding: 10px; display: flex; flex-direction: column; justify-content: center;">
                <div style="text-align:center">
                    <img src="{time}" alt="Project Image" style="width:80px;height:auto;text-align:center">
				</div>
                <p>Period: {portfolio.start_date} - {portfolio.end_date}</p>
            </div>
        </div>
		<div style="display: flex;">
            <div style="display: flex; flex-direction: column; width:40%">
                <div style="width: 100%;border: 1px solid gray; height:100px;">
                    <img src="{client_logo}" style="width:100%;height:100%;object-fit:contain;"/>
				</div>
                <div style="width: 100%;border: 1px solid gray;height:100px;">
                    Client reference: {client_reference}
				</div>
                <div style="width: 100%;border: 1px solid gray;height:100px;">
                    Client contact: {client_contact}
				</div>
            </div>
            <div style="flex: 1;width:60%; height: 300px;border: 1px solid gray;">
                {images_list}
			</div>
        </div>
		
		<div style="display: flex; width: 100%;">
            <div style="flex: 1; margin: 0;  border: 1px solid gray; padding: 10px; display: flex; flex-direction: column; width:60%">
			    <p>Project Description</p>
                <p>{portfolio.body}</p>
            </div>
            <div style="flex: 1; margin: 0; border: 1px solid gray; padding: 10px; display: flex; flex-direction: column; width:40%">
                <p>Services Provided</p>
				<ul>
				{services_list}
				</ul>
            </div>
        </div>
		
		<div>
		    <img src="{footer}" alt="Project Image" style="width:100%;height:auto;text-align:center;postion:absolute;bottom:-1px;left:0px">
		</div>
        """

    content = f"""
    <html>
    <head>
        <title>Kartoza Project Sheet</title>
    </head>
    <body>
        
        {project_details}
    </body>
    </html>
    """
    return content


def generate_docx(html_content):
    doc = Document()
    soup = BeautifulSoup(html_content, 'html.parser')

    # Handling headers and paragraphs with styles
    for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'li', 'img', 'div']):
        if element.name.startswith('h'):
            doc.add_heading(element.get_text(), level=int(element.name[1]))
        elif element.name == 'p':
            p = doc.add_paragraph(element.get_text())
            if 'style' in element.attrs:
                style = element.attrs['style']
                if 'text-align:center' in style:
                    p.alignment = 1  # Center alignment
        elif element.name == 'ul':
            for li in element.find_all('li'):
                doc.add_paragraph(f'• {li.get_text()}', style='ListBullet')
        elif element.name == 'img':
            response = requests.get(element['src'])
            img = io.BytesIO(response.content)
            doc.add_picture(img, width=Inches(2.0))  # Adjust size as needed
        elif element.name == 'div':
            style = element.attrs.get('style', '')
            if 'display: flex' in style:
                # Handle flex layout divs as tables
                divs = element.find_all('div', recursive=False)
                if not divs:
                    continue  # Skip if no divs are found
                num_columns = len(divs[0].find_all('div', recursive=False))
                table = doc.add_table(rows=0, cols=num_columns)
                
                for row_div in divs:
                    cells = row_div.find_all('div', recursive=False)
                    row = table.add_row().cells
                    for i, cell in enumerate(cells):
                        if i < len(row):  # Ensure not to exceed the number of columns
                            p = row[i].add_paragraph(cell.get_text())
                            if 'text-align:center' in style:
                                p.alignment = 1  # Center alignment

            # Additional handling for specific divs
            if 'width:40%' in style:
                # Special handling for specific divs like the one with width:40%
                pass
            elif 'width:60%' in style:
                # Special handling for specific divs like the one with width:60%
                pass
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def worldbank_format(portfolios):
    """Create a World Bank format document for the given portfolios."""
    portfolio_names = frappe.parse_json(portfolios)
    doc = Document()

    # Add Title
    title = doc.add_heading(level=1)
    title_run = title.add_run("Assignment Details")
    title_run.bold = True

    # Loop through each portfolio and create a table
    for portfolio_name in portfolio_names:
        details = frappe.get_doc("Portfolio", portfolio_name)

        # Add a heading for each portfolio
        doc.add_heading(details.title, level=2)

        # Create a table for the details
        table = doc.add_table(rows=14, cols=2)
        table.style = 'Table Grid'
        table.autofit = False

        # Set the width of the table columns
        for row in table.rows:
            row.cells[0].width = Pt(200)
            row.cells[1].width = Pt(300)

        # Add the details to the table
        details_dict = {
            "Assignment name:": details.title,
            "Approx. value of the contract (in current US$):": details.approximate_contract_value,
            "Country:": details.location,
            "Duration of assignment (months):": details.duration_of_assignment,
            "Name of Client(s):": details.client,
            "Contact Person, Title/Designation, Tel. No./Address:": details.contact,
            "Start Date (month/year):": details.start_date,
            "End Date (month/year):": details.end_date,
            "Total No. of staff-months of the assignment:": details.total_staff_months,
            "No. of professional staff-months provided by your consulting firm/organization or "
            "your sub consultants:": details.total_staff_months,
            "Name of associated Consultants, if any:": "",
            "Name of senior professional staff of your consulting firm/organization involved and "
            "designation and/or functions performed (e.g. Project Director/ Coordinator, "
            "Team Leader):": "",
            "Description of Project:": details.body,
            "Description of actual services provided by your staff within the "
            "assignment:": details.services_listed,
        }

        # Populate the table with the correct details
        for i, (key, value) in enumerate(details_dict.items()):
            cell1 = table.cell(i, 0)
            cell2 = table.cell(i, 1)
            cell1.text = key
            cell2.text = str(value) if value else ""

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()
