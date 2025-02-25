import streamlit as st
import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from werkzeug.utils import secure_filename
import tempfile
import shutil
from streamlit_quill import st_quill

TEMPLATE_PATH = 'workshop_template.docx'  # Path to your report template

def add_underline(run):
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    run._element.rPr.append(u)

def add_centered_image(doc, image_path, width=Cm(9), height=Cm(16)):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    run.add_picture(image_path, width=width, height=height)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

def ensure_heading_style(doc, style_name, font_name='DIN Pro Regular', font_size=12, rgb_color=(135, 206, 235)):
    styles = doc.styles
    try:
        style = styles[style_name]
    except KeyError:
        style = styles.add_style(style_name, 1)
    
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)
    font.color.rgb = RGBColor(*rgb_color)
    return style

def save_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    
    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, secure_filename(uploaded_file.name))
    
    with open(temp_path, 'wb') as f:
        f.write(uploaded_file.getvalue())
    
    return temp_path

def add_signature_section(doc, faculty_name, hod_name):
    # Add space before signatures
    doc.add_paragraph('\n\n')
    
    # Create a table for signatures with specific widths
    signature_table = doc.add_table(rows=2, cols=2)
    signature_table.autofit = False
    
    # Set column widths (50% each)
    for cell in signature_table.columns[0].cells:
        cell.width = Cm(8)
    for cell in signature_table.columns[1].cells:
        cell.width = Cm(8)
    
    # Faculty-in-charge section
    faculty_label_cell = signature_table.cell(0, 0)
    faculty_label_paragraph = faculty_label_cell.paragraphs[0]
    faculty_label_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    faculty_label_run = faculty_label_paragraph.add_run("Name & Signature of the")
    faculty_label_run.font.name = 'Times New Roman'
    faculty_label_run.font.size = Pt(11)
    
    faculty_label_cell2 = signature_table.cell(1, 0)
    faculty_label_paragraph2 = faculty_label_cell2.paragraphs[0]
    faculty_label_paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    faculty_label_run2 = faculty_label_paragraph2.add_run("Faculty-in-charge")
    faculty_label_run2.font.name = 'Times New Roman'
    faculty_label_run2.font.size = Pt(11)
    faculty_name_run = faculty_label_paragraph2.add_run(f"\n{faculty_name}")
    faculty_name_run.font.name = 'Times New Roman'
    faculty_name_run.font.size = Pt(11)
    
    # HoD section
    hod_label_cell = signature_table.cell(0, 1)
    hod_label_paragraph = hod_label_cell.paragraphs[0]
    hod_label_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    hod_label_run = hod_label_paragraph.add_run("Name & Signature of")
    hod_label_run.font.name = 'Times New Roman'
    hod_label_run.font.size = Pt(11)
    
    hod_label_cell2 = signature_table.cell(1, 1)
    hod_label_paragraph2 = hod_label_cell2.paragraphs[0]
    hod_label_paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    hod_label_run2 = hod_label_paragraph2.add_run("HoD")
    hod_label_run2.font.name = 'Times New Roman'
    hod_label_run2.font.size = Pt(11)
    hod_name_run = hod_label_paragraph2.add_run(f"\n{hod_name}")
    hod_name_run.font.name = 'Times New Roman'
    hod_name_run.font.size = Pt(11)

def create_report(inputs, files):
    temp_dir = tempfile.mkdtemp()
    try:
        doc = Document(TEMPLATE_PATH)
        
        # Set up styles
        ensure_heading_style(doc, 'Heading 1')
        ensure_heading_style(doc, 'Heading 2')
        
        # Add department name
        header = doc.add_paragraph()
        run = header.add_run(f"Department of {inputs['department_name']}")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(79, 129, 189)
        add_underline(run)
        header.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        
        # Add title
        title = doc.add_paragraph()
        run = title.add_run("Master Class")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(79, 129, 189)
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc.add_paragraph('\n')
        
        # Create details table
        table = doc.add_table(rows=7, cols=2)
        table.style = 'Table Grid'
        
        details = [
            ("Topic", inputs['topic']),
            ("Expert", inputs['expert']),
            ("Venue", inputs['venue']),
            ("Date", inputs['date']),
            ("Time", inputs['time']),
            ("Faculty Coordinator", inputs['coordinator']),
            ("Number of Participants", inputs['num_participants'])
        ]
        
        for i, (label, value) in enumerate(details):
            row = table.rows[i]
            label_cell = row.cells[0]
            label_paragraph = label_cell.paragraphs[0]
            label_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            label_run = label_paragraph.add_run(label)
            label_run.font.name = 'Times New Roman'
            label_run.font.size = Pt(11)
            label_run.bold = True

            value_cell = row.cells[1]
            value_paragraph = value_cell.paragraphs[0]
            value_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            value_run = value_paragraph.add_run(str(value))
            value_run.font.name = 'Times New Roman'
            value_run.font.size = Pt(11)
        
        # Add sections
        sections = [
            ("Summary of the Event:", inputs['summary']),
            ("Outcome of the Event:", inputs['outcome'])
        ]
        
        for heading_text, content in sections:
            doc.add_paragraph('\n')
            heading = doc.add_paragraph()
            run = heading.add_run(heading_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(79, 129, 189)
            
            content_para = doc.add_paragraph(content)
            content_run = content_para.runs[0]
            content_run.font.name = 'Times New Roman'
            content_run.font.size = Pt(11)
        
        # Add images sections
        if files['invite_image']:
            doc.add_page_break()
            invite_heading = doc.add_paragraph()
            run = invite_heading.add_run("Invite")
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(79, 129, 189)
            invite_heading.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
            invite_path = save_uploaded_file(files['invite_image'])
            if invite_path:
                add_centered_image(doc, invite_path, width=Cm(9), height=Cm(17))
        
        image_sections = [
            ("Action Photos", files['action_photos'], Cm(10), Cm(10)),
            ("Attendance Sheet", files['attendance_photos'], Cm(9), Cm(9)),
            ("Analysis Report", files['analysis_photos'], Cm(9), Cm(9))
        ]
        
        for heading_text, photos, width, height in image_sections:
            if photos:
                doc.add_page_break()
                heading = doc.add_paragraph()
                run = heading.add_run(heading_text)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(79, 129, 189)
                heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                doc.add_paragraph('\n')
                
                for photo in photos:
                    photo_path = save_uploaded_file(photo)
                    if photo_path:
                        add_centered_image(doc, photo_path, width=width, height=height)
        
        # Add signature section with fixed font
        add_signature_section(doc, inputs['faculty_in_charge'], inputs['hod_name'])
        
        # Save document
        output_path = os.path.join(temp_dir, 'Workshop_Report.docx')
        doc.save(output_path)
        
        with open(output_path, 'rb') as f:
            file_data = f.read()
        
        return file_data
        
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            st.warning(f"Could not clean up temporary files: {str(e)}")

def main():
    st.title("Workshop Report Generator")
    
    # Input fields
    department_name = st.text_input("Department Name")
    topic = st.text_input("Topic")
    expert = st.text_input("Expert Name")
    venue = st.text_input("Venue")
    date = st.date_input("Date")
    time = st.time_input("Time")
    coordinator = st.text_input("Faculty Coordinator")
    num_participants = st.number_input("Number of Participants", min_value=1)
    summary = st.text_area("Summary of the Event")
    outcome = st.text_area("Outcome of the Event")
    faculty_in_charge = st.text_input("Name of the Faculty-in-charge")
    hod_name = st.text_input("Name of the HoD")
    
    # File uploads
    invite_image = st.file_uploader("Upload Invite Image", type=['png', 'jpg', 'jpeg'])
    action_photos = st.file_uploader("Upload Action Photos", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    attendance_photos = st.file_uploader("Upload Attendance Sheet Photos", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    analysis_photos = st.file_uploader("Upload Analysis Report Photos", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
    
    if st.button("Generate Report"):
        if not all([department_name, topic, expert, venue, coordinator, summary, outcome]):
            st.error("Please fill in all required fields")
            return
        
        inputs = {
            'department_name': department_name,
            'topic': topic,
            'expert': expert,
            'venue': venue,
            'date': date.strftime('%d-%m-%Y'),
            'time': time.strftime('%H:%M'),
            'coordinator': coordinator,
            'num_participants': num_participants,
            'summary': summary,
            'outcome': outcome,
            'faculty_in_charge': faculty_in_charge,
            'hod_name': hod_name
        }
        
        files = {
            'invite_image': invite_image,
            'action_photos': action_photos,
            'attendance_photos': attendance_photos,
            'analysis_photos': analysis_photos
        }
        
        try:
            file_data = create_report(inputs, files)
            
            st.download_button(
                label="Download Report",
                data=file_data,
                file_name="Workshop_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("Report generated successfully!")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.warning("Please check the inputs and try again.")

if __name__ == "__main__":
    main()
